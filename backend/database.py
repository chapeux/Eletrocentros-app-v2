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
            motivo TEXT,
            anexo_nome VARCHAR(255),
            anexo_caminho TEXT,
            INDEX idx_data_hora (data_hora),
            INDEX idx_grupo_campo (grupo_area, regra_campo)
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
        """
        cursor.execute(create_table_sql)

        # 3. Garante que a tabela 'app_settings' exista para centralização de dados
        create_settings_table_sql = r"""
        CREATE TABLE IF NOT EXISTS app_settings (
            chave VARCHAR(64) PRIMARY KEY,
            valor LONGTEXT NOT NULL,
            versao INT NOT NULL DEFAULT 1,
            atualizado_em DATETIME NOT NULL,
            atualizado_por VARCHAR(100) NOT NULL
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
        """
        cursor.execute(create_settings_table_sql)

        # 4. Garante que a tabela 'app_releases' exista para distribuição de pacotes via banco
        create_releases_table_sql = r"""
        CREATE TABLE IF NOT EXISTS app_releases (
            id              INT AUTO_INCREMENT PRIMARY KEY,
            version         VARCHAR(20)  NOT NULL UNIQUE,
            descricao       TEXT,
            obrigatoria     BOOLEAN      NOT NULL DEFAULT FALSE,
            pacote_zip      LONGBLOB     NOT NULL,
            sha256          CHAR(64)     NOT NULL,
            tamanho_bytes   BIGINT       NOT NULL,
            status          ENUM('rascunho','publicada','revogada') NOT NULL DEFAULT 'rascunho',
            publicado_em    DATETIME     NOT NULL,
            publicado_por   VARCHAR(100) NOT NULL,
            INDEX idx_status (status)
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
        """
        cursor.execute(create_releases_table_sql)

        # 5. Garante adição das colunas 'motivo', 'anexo_nome' e 'anexo_caminho' em logs_modificacoes
        try:
            cursor.execute("SHOW COLUMNS FROM logs_modificacoes LIKE 'motivo';")
            if not cursor.fetchone():
                cursor.execute("ALTER TABLE logs_modificacoes ADD COLUMN motivo TEXT;")
            
            cursor.execute("SHOW COLUMNS FROM logs_modificacoes LIKE 'anexo_nome';")
            if not cursor.fetchone():
                cursor.execute("ALTER TABLE logs_modificacoes ADD COLUMN anexo_nome VARCHAR(255);")

            cursor.execute("SHOW COLUMNS FROM logs_modificacoes LIKE 'anexo_caminho';")
            if not cursor.fetchone():
                cursor.execute("ALTER TABLE logs_modificacoes ADD COLUMN anexo_caminho TEXT;")
        except Exception as ex:
            print(f"[Database Python] Aviso ao verificar colunas adicionais: {ex}")

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
    usuario: Optional[str] = None,
    motivo: Optional[str] = None,
    anexo_nome: Optional[str] = None,
    anexo_caminho: Optional[str] = None
) -> bool:
    """Insere um novo registro na tabela 'logs_modificacoes' com anexo opcional."""
    if not usuario:
        usuario = get_current_user()

    data_hora = datetime.now()
    antes_str = json.dumps(antes, ensure_ascii=False) if antes is not None else None
    depois_str = json.dumps(depois, ensure_ascii=False) if depois is not None else None

    sql = f"""
        INSERT INTO {DB_TABELA} 
        (data_hora, usuario, grupo_area, regra_campo, subtab, antes, depois, detalhes, motivo, anexo_nome, anexo_caminho)
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
    """
    valores = (data_hora, usuario, grupo_area, regra_campo, subtab, antes_str, depois_str, detalhes, motivo, anexo_nome, anexo_caminho)

    try:
        conn = get_db_connection(with_db=True)
        cursor = conn.cursor()
        cursor.execute(sql, valores)
        conn.commit()
        cursor.close()
        conn.close()
        anexo_info = f" (Anexo: {anexo_nome})" if anexo_nome else ""
        print(f"[Database Python] Log gravado em {DB_TABELA}: [{grupo_area} -> {regra_campo} ({subtab})] por '{usuario}' (Motivo: {motivo}){anexo_info}")
        return True
    except Exception as e:
        print(f"[Database Python] Erro ao gravar log no MySQL: {e}")
        return False


def comparar_e_registrar_alteracoes(
    regras_antigas: List[Dict],
    regras_novas: List[Dict],
    usuario: Optional[str] = None,
    motivo: Optional[str] = None,
    anexo_nome: Optional[str] = None,
    anexo_caminho: Optional[str] = None
) -> int:
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
                        usuario=usuario,
                        motivo=motivo,
                        anexo_nome=anexo_nome,
                        anexo_caminho=anexo_caminho
                    )
                    if sucesso:
                        total_logs += 1

    return total_logs


def obter_logs(limit: int = 100) -> List[Dict[str, Any]]:
    """Retorna os logs de alterações mais recentes cadastrados na tabela 'logs_modificacoes'."""
    sql = f"""
        SELECT id, data_hora, usuario, grupo_area, regra_campo, subtab, antes, depois, detalhes, motivo, anexo_nome, anexo_caminho
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

            # Processa lista de anexos (suporta múltiplos ou anexo único legado)
            anexos_lista = []
            caminho_raw = r.get("anexo_caminho")
            nome_raw = r.get("anexo_nome")

            if caminho_raw:
                caminho_str = str(caminho_raw).strip()
                if caminho_str.startswith("[") and caminho_str.endswith("]"):
                    try:
                        parsed = json.loads(caminho_str)
                        if isinstance(parsed, list):
                            anexos_lista = parsed
                    except Exception:
                        pass
                if not anexos_lista:
                    anexos_lista = [{"nome": nome_raw or os.path.basename(caminho_str), "caminho": caminho_str}]

            r["anexos"] = anexos_lista
            logs.append(r)
        cursor.close()
        conn.close()
    except Exception as e:
        print(f"[Database Python] Erro ao buscar logs em {DB_TABELA}: {e}")

    return logs
