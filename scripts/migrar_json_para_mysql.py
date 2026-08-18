"""
Script de Migração Única: Arquivos JSON para Tabela 'app_settings' no MySQL.
Carrega 'config.json', 'regras.json', 'seletor.json' e 'template_blocks.json'
na base central bd_eletrocentros_app.app_settings.
"""

import json
import sys
from datetime import datetime
from pathlib import Path

# Adiciona a raiz do projeto ao path para importar o módulo backend
PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from backend.database import get_db_connection, init_db

FRONTEND_DIR = PROJECT_ROOT / "frontend"

ARQUIVOS_MIGRACAO = {
    "config_geral":    FRONTEND_DIR / "config.json",
    "regras_calculo":  FRONTEND_DIR / "regras.json",
    "seletor_pep_cts": FRONTEND_DIR / "seletor.json",
    "template_blocks": FRONTEND_DIR / "template_blocks.json",
}


def migrar():
    print("=" * 65)
    print(" [MIGRAÇÃO] Iniciando migração dos arquivos JSON para o MySQL")
    print(f" Servidor : dcprd036187.weg.net:8502")
    print(f" Database : bd_eletrocentros_app")
    print(f" Tabela   : app_settings")
    print("=" * 65)

    # 1. Garante que banco e tabelas existam
    if not init_db():
        print("[ERRO] Falha ao inicializar banco de dados MySQL. Abortando.")
        return False

    conn = get_db_connection(with_db=True)
    cursor = conn.cursor()

    total_migrado = 0

    for chave, caminho in ARQUIVOS_MIGRACAO.items():
        if not caminho.exists():
            print(f"⚠️ [AVISO] Arquivo não encontrado: {caminho} (Pulando)")
            continue

        try:
            texto = caminho.read_text(encoding="utf-8")
            dados = json.loads(texto)
            tamanho_kb = len(texto.encode("utf-8")) / 1024.0
            agora = datetime.now()

            # Insere ou atualiza mantendo integridade
            sql = """
            INSERT INTO app_settings (chave, valor, versao, atualizado_em, atualizado_por)
            VALUES (%s, %s, 1, %s, 'migracao_inicial')
            ON DUPLICATE KEY UPDATE
                valor = VALUES(valor),
                atualizado_em = VALUES(atualizado_em),
                atualizado_por = 'migracao_inicial';
            """
            cursor.execute(sql, (chave, json.dumps(dados, ensure_ascii=False), agora))
            conn.commit()

            print(f"  [OK] {chave:18} <- {caminho.name:22} ({tamanho_kb:6.2f} KB) migrado com sucesso.")
            total_migrado += 1

        except Exception as e:
            print(f"  [ERRO] Falha ao migrar {chave} de {caminho}: {e}")

    cursor.close()
    conn.close()

    print("=" * 65)
    print(f" [MIGRAÇÃO CONCLUÍDA] {total_migrado} de {len(ARQUIVOS_MIGRACAO)} chaves salvas no MySQL!")
    print("=" * 65)
    return True


if __name__ == "__main__":
    sucesso = migrar()
    sys.exit(0 if sucesso else 1)
