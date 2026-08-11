"""
Módulo de Banco de Dados MySQL — Eletrocentros App
Gerencia a conexão com o MySQL, auto-criação do BD/tabela 'logs_modificacoes' e o registro de histórico de alterações de regras.
"""

import os
import json
import getpass
from datetime import datetime
from typing import List, Dict, Any, Optional

try:
    import mysql.connector
    HAS_MYSQL = True
except ImportError:
    HAS_MYSQL = False

# ─── CONFIGURAÇÃO DO BANCO DE DADOS ──────────────────────────────────────────
DB_CONFIG = {
    "host":     "dcprd036187.weg.net",
    "database": "bd_eletrocentros_app",  
    "user":     "root",
    "password": "neoempresarial",
    "port":     8502,
}
DB_TABELA = "logs_modificacoes"


def get_current_user() -> str:
    """Retorna o nome do usuário atual do sistema operacional (ou variável de ambiente)."""
    try:
        user = getpass.getuser()
        if user:
            return user
    except Exception:
        pass
    return os.environ.get("USERNAME", os.environ.get("USER", "usuario_sistema"))


def get_db_connection(with_db: bool = True):
    """Retorna uma conexão MySQL ativa. Se with_db=False, conecta sem selecionar o banco de dados."""
    if not HAS_MYSQL:
        raise ImportError("O módulo 'mysql.connector' não está instalado no ambiente Python.")

    config = dict(DB_CONFIG)
    if not with_db:
        config.pop("database", None)

    return mysql.connector.connect(**config)


def init_db() -> bool:
    """
    Garante que o banco de dados 'bd_eletrocentros_app' e a tabela 'logs_modificacoes'
    existam no servidor MySQL com a codificação utf8mb4.
    """
    if not HAS_MYSQL:
        print("[Database Python] AVISO: 'mysql.connector' não encontrado. Logs de BD estarão desativados.")
        return False

    try:
        # 1. Garantir que o banco de dados exista
        conn = get_db_connection(with_db=False)
        cursor = conn.cursor()
        cursor.execute("CREATE DATABASE IF NOT EXISTS bd_eletrocentros_app CHARACTER SET utf8mb4;")
        cursor.close()
        conn.close()

        # 2. Garantir que a tabela 'logs_modificacoes' exista
        conn = get_db_connection(with_db=True)
        cursor = conn.cursor()
        create_table_sql = r"""
        CREATE TABLE IF NOT EXISTS logs_modificacoes (
            id INT AUTO_INCREMENT PRIMARY KEY,
            data_hora DATETIME NOT NULL,
            usuario VARCHAR(100) NOT NULL,
            grupo_area VARCHAR(150) NOT NULL,
            regra_campo VARCHAR(150) NOT NULL,
            subtab VARCHAR(20) NOT NULL,
            antes LONGTEXT,
            depois LONGTEXT,
            detalhes TEXT,
            INDEX idx_data_hora (data_hora),
            INDEX idx_grupo_campo (grupo_area, regra_campo)
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
        """
        cursor.execute(create_table_sql)
        conn.commit()
        cursor.close()
        conn.close()

        print(f"[Database Python] Conexão MySQL com '{DB_CONFIG['database']}.{DB_TABELA}' verificada e inicializada com sucesso!")
        return True
    except Exception as e:
        print(f"[Database Python] Erro ao inicializar BD MySQL: {e}")
        return False


def registrar_log(
    grupo_area: str,
    regra_campo: str,
    subtab: str,
    antes: Any,
    depois: Any,
    detalhes: str = "",
    usuario: Optional[str] = None
) -> bool:
    """Insere um novo registro na tabela 'logs_modificacoes'."""
    if not usuario:
        usuario = get_current_user()

    data_hora = datetime.now()
    antes_str = json.dumps(antes, ensure_ascii=False) if antes is not None else None
    depois_str = json.dumps(depois, ensure_ascii=False) if depois is not None else None

    sql = f"""
        INSERT INTO {DB_TABELA} 
        (data_hora, usuario, grupo_area, regra_campo, subtab, antes, depois, detalhes)
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
    """
    valores = (data_hora, usuario, grupo_area, regra_campo, subtab, antes_str, depois_str, detalhes)

    try:
        conn = get_db_connection(with_db=True)
        cursor = conn.cursor()
        cursor.execute(sql, valores)
        conn.commit()
        cursor.close()
        conn.close()
        print(f"[Database Python] Log gravado em {DB_TABELA}: [{grupo_area} -> {regra_campo} ({subtab})] por '{usuario}'")
        return True
    except Exception as e:
        print(f"[Database Python] Erro ao gravar log no MySQL: {e}")
        return False


def comparar_e_registrar_alteracoes(regras_antigas: List[Dict], regras_novas: List[Dict], usuario: Optional[str] = None) -> int:
    """
    Compara o estado anterior e novo das regras (estrutura de regras.json)
    e registra um log para cada sub-regra (H ou DUR) que sofreu alteração.
    Retorna a quantidade de alterações registradas.
    """
    if not usuario:
        usuario = get_current_user()

    # Mapeia regras antigas por área e por campo para busca rápida
    map_antigas = {}
    for area_obj in (regras_antigas or []):
        area_nome = area_obj.get("area", "")
        map_antigas[area_nome] = area_obj.get("campos", {})

    total_logs = 0

    for area_obj in (regras_novas or []):
        area_nome = area_obj.get("area", "")
        campos_novos = area_obj.get("campos", {})
        campos_antigos = map_antigas.get(area_nome, {})

        for campo_key, campo_novo_data in campos_novos.items():
            campo_antigo_data = campos_antigos.get(campo_key, {})

            for subtab in ["H", "DUR"]:
                regra_antiga = campo_antigo_data.get(subtab)
                regra_nova = campo_novo_data.get(subtab)

                # Se a sub-regra existe no novo objeto ou mudou de valor
                if json.dumps(regra_antiga, sort_keys=True) != json.dumps(regra_nova, sort_keys=True):
                    detalhes = f"Alteração da regra de {subtab} do campo '{campo_key}' na disciplina/área '{area_nome}'."
                    sucesso = registrar_log(
                        grupo_area=area_nome,
                        regra_campo=campo_key,
                        subtab=subtab,
                        antes=regra_antiga,
                        depois=regra_nova,
                        detalhes=detalhes,
                        usuario=usuario
                    )
                    if sucesso:
                        total_logs += 1

    return total_logs


def obter_logs(limit: int = 100) -> List[Dict[str, Any]]:
    """Retorna os logs de alterações mais recentes cadastrados na tabela 'logs_modificacoes'."""
    sql = f"""
        SELECT id, data_hora, usuario, grupo_area, regra_campo, subtab, antes, depois, detalhes
        FROM {DB_TABELA}
        ORDER BY id DESC
        LIMIT %s
    """
    logs = []
    try:
        conn = get_db_connection(with_db=True)
        cursor = conn.cursor(dictionary=True)
        cursor.execute(sql, (limit,))
        rows = cursor.fetchall()
        for r in rows:
            # Formata data_hora para string ISO/BR se for datetime
            if isinstance(r.get("data_hora"), datetime):
                r["data_hora"] = r["data_hora"].strftime("%d/%m/%Y %H:%M:%S")
            # Tenta converter os campos de JSON string de volta para objeto se necessário
            for key in ["antes", "depois"]:
                if isinstance(r.get(key), str):
                    try:
                        r[key] = json.loads(r[key])
                    except Exception:
                        pass
            logs.append(r)
        cursor.close()
        conn.close()
    except Exception as e:
        print(f"[Database Python] Erro ao buscar logs em {DB_TABELA}: {e}")

    return logs
