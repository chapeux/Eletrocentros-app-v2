"""
Módulo de Armazenamento de Configurações e Regras no MySQL — Eletrocentros App
Substitui a leitura/escrita direta de arquivos .json locais por persistência centralizada
na tabela 'app_settings' com controle de concorrência (lock otimista) e fallback resiliente.
"""

import json
import os
import shutil
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Optional, Tuple

from backend.database import get_current_user, get_db_connection, HAS_MYSQL


def _salvar_arquivo_atomico(caminho: Path, conteudo_dict: Dict[str, Any]) -> bool:
    """Grava o arquivo JSON de backup local de forma atômica (usando .tmp e replace)."""
    try:
        caminho.parent.mkdir(parents=True, exist_ok=True)
        tmp_file = caminho.with_suffix(".tmp")
        with open(tmp_file, "w", encoding="utf-8") as f:
            json.dump(conteudo_dict, f, ensure_ascii=False, indent=2)
        os.replace(tmp_file, caminho)
        return True
    except Exception as e:
        print(f"[SettingsStore] Aviso ao gravar cache local '{caminho}': {e}")
        return False


def get_setting(chave: str, fallback_file: Optional[Path] = None) -> Tuple[Dict[str, Any], int]:
    """
    Retorna (valor_dict, versao) da configuração no MySQL.
    
    Retornos de versão:
      versao > 0 : Registro lido com sucesso do MySQL.
      versao = 0 : Registro não existia no MySQL (carregado do fallback local).
      versao = -1: Modo offline (falha de conexão com MySQL, lido do fallback).
    """
    if HAS_MYSQL:
        try:
            conn = get_db_connection(with_db=True)
            cursor = conn.cursor()
            cursor.execute("SELECT valor, versao FROM app_settings WHERE chave = %s", (chave,))
            row = cursor.fetchone()
            cursor.close()
            conn.close()

            if row is not None:
                valor_raw, versao = row
                if isinstance(valor_raw, str):
                    dados = json.loads(valor_raw)
                else:
                    dados = valor_raw or {}
                
                # Atualiza silenciosamente o cache local se fallback_file for fornecido
                if fallback_file and dados:
                    _salvar_arquivo_atomico(fallback_file, dados)

                return dados, int(versao)
        except Exception as e:
            print(f"[SettingsStore] Erro ao conectar ao MySQL para chave '{chave}': {e}. Acionando fallback...")

    # Fallback local se MySQL estiver indisponível ou se o registro for novo
    if fallback_file and fallback_file.exists():
        try:
            dados = json.loads(fallback_file.read_text(encoding="utf-8"))
            # Se HAS_MYSQL falhou, marca como -1 (offline)
            versao_fallback = -1 if HAS_MYSQL else 0
            return dados, versao_fallback
        except Exception as e:
            print(f"[SettingsStore] Erro ao ler arquivo de fallback '{fallback_file}': {e}")

    return {}, 0


def save_setting(
    chave: str,
    valor: Dict[str, Any],
    versao_esperada: int,
    usuario: Optional[str] = None,
    backup_file: Optional[Path] = None
) -> Dict[str, Any]:
    """
    Salva configuração no MySQL com lock otimista:
    Só grava se versao_esperada bater com a versão atual no banco.
    """
    if versao_esperada == -1:
        return {
            "status": "error",
            "message": "Sem conexão com o banco de dados MySQL. Não é possível salvar alterações em modo offline."
        }

    usuario = usuario or get_current_user()
    valor_json = json.dumps(valor, ensure_ascii=False)
    agora = datetime.now()

    if not HAS_MYSQL:
        # Modo sem driver MySQL: salva apenas no arquivo local
        if backup_file:
            if _salvar_arquivo_atomico(backup_file, valor):
                return {"status": "success", "versao": versao_esperada + 1}
        return {"status": "error", "message": "Módulo 'mysql.connector' não instalado."}

    try:
        conn = get_db_connection(with_db=True)
        cursor = conn.cursor()

        if versao_esperada == 0:
            # Primeira gravação / inserção inicial
            cursor.execute(
                """INSERT INTO app_settings (chave, valor, versao, atualizado_em, atualizado_por)
                   VALUES (%s, %s, 1, %s, %s)
                   ON DUPLICATE KEY UPDATE
                       valor = VALUES(valor),
                       versao = versao + 1,
                       atualizado_em = VALUES(atualizado_em),
                       atualizado_por = VALUES(atualizado_por)""",
                (chave, valor_json, agora, usuario)
            )
            conn.commit()
            
            # Obtém a nova versão resultante
            cursor.execute("SELECT versao FROM app_settings WHERE chave = %s", (chave,))
            row = cursor.fetchone()
            nova_versao = int(row[0]) if row else 1
            cursor.close()
            conn.close()

            if backup_file:
                _salvar_arquivo_atomico(backup_file, valor)

            return {"status": "success", "versao": nova_versao}

        # Gravação normal com lock otimista
        cursor.execute(
            """UPDATE app_settings
               SET valor = %s, versao = versao + 1, atualizado_em = %s, atualizado_por = %s
               WHERE chave = %s AND versao = %s""",
            (valor_json, agora, usuario, chave, versao_esperada)
        )
        conn.commit()
        afetadas = cursor.rowcount
        
        nova_versao = versao_esperada + 1
        if afetadas == 0:
            # Houve conflito de concorrência — alguém salvou antes
            cursor.execute("SELECT versao, atualizado_por, atualizado_em FROM app_settings WHERE chave = %s", (chave,))
            conflito_row = cursor.fetchone()
            cursor.close()
            conn.close()

            detalhes_conflito = ""
            if conflito_row:
                v_atual, u_atual, d_atual = conflito_row
                detalhes_conflito = f" (Versão atual: {v_atual}, salva por '{u_atual}' em {d_atual.strftime('%d/%m/%Y %H:%M:%S') if isinstance(d_atual, datetime) else d_atual})"

            return {
                "status": "conflito",
                "message": f"Alguém salvou uma versão mais recente das regras enquanto você editava{detalhes_conflito}. Por favor, recarregue os dados e refaça a alteração."
            }

        cursor.close()
        conn.close()

        # Atualiza o cache local
        if backup_file:
            _salvar_arquivo_atomico(backup_file, valor)

        return {"status": "success", "versao": nova_versao}

    except Exception as e:
        print(f"[SettingsStore] Erro ao gravar chave '{chave}' no MySQL: {e}")
        return {
            "status": "error",
            "message": f"Erro de banco de dados ao salvar: {str(e)}"
        }
