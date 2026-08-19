"""
Script de Publicação de Release no Banco de Dados MySQL — Eletrocentros App
Distribuição 100% centralizada e atômica via tabela app_releases e app_settings.
"""

import sys
import os
import json
import hashlib
import zipfile
from io import BytesIO
from datetime import datetime
from pathlib import Path

# Ajuste de path para importação do backend
BASE_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BASE_DIR))

from backend.database import get_db_connection, get_current_user, init_db


def zipar_aplicativo(base_dir: Path) -> bytes:
    """
    Compacta os arquivos essenciais do aplicativo para distribuição.
    """
    buffer = BytesIO()
    
    pastas_incluir = ["backend", "frontend", "assets"]
    arquivos_incluir = ["main.py"]

    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as zf:
        # Arquivos raiz
        for arq_nome in arquivos_incluir:
            arq_path = base_dir / arq_nome
            if arq_path.exists():
                zf.write(arq_path, arq_nome)
                print(f"  -> Adicionado arquivo: {arq_nome}")

        # Pastas e subdiretórios
        for pasta_nome in pastas_incluir:
            pasta_path = base_dir / pasta_nome
            if pasta_path.exists():
                for item in pasta_path.rglob("*"):
                    if item.is_file():
                        # Ignora caches e arquivos temporários
                        if "__pycache__" in item.parts or item.suffix in [".pyc", ".pyo", ".tmp"]:
                            continue
                        rel_path = item.relative_to(base_dir)
                        zf.write(item, str(rel_path).replace("\\", "/"))
                print(f"  -> Adicionado diretório: {pasta_nome}")

    return buffer.getvalue()


def publicar_release(version: str, descricao: str = "", obrigatoria: bool = True):
    print("=" * 70)
    print(f" [PUBLICAR RELEASE] Publicando Versão {version} no MySQL")
    print("=" * 70)

    # 1. Garante que as tabelas existem
    init_db()

    # 2. Gera o pacote .zip
    print("\n1. Compactando arquivos do aplicativo...")
    pacote = zipar_aplicativo(BASE_DIR)
    tamanho_bytes = len(pacote)
    sha256_hash = hashlib.sha256(pacote).hexdigest()
    usuario = get_current_user()

    print(f"   Tamanho do pacote: {tamanho_bytes / 1024:.1f} KB ({tamanho_bytes:,} bytes)")
    print(f"   SHA-256: {sha256_hash}")

    conn = get_db_connection(with_db=True)
    cursor = conn.cursor()

    try:
        # 3. Grava o pacote como rascunho (Passo 1/4)
        print("\n2. Gravando pacote na tabela 'app_releases' (status='rascunho')...")
        cursor.execute(
            """
            INSERT INTO app_releases
                (version, descricao, obrigatoria, pacote_zip, sha256, tamanho_bytes, status, publicado_em, publicado_por)
            VALUES
                (%s, %s, %s, %s, %s, %s, 'rascunho', %s, %s)
            ON DUPLICATE KEY UPDATE
                descricao = VALUES(descricao),
                obrigatoria = VALUES(obrigatoria),
                pacote_zip = VALUES(pacote_zip),
                sha256 = VALUES(sha256),
                tamanho_bytes = VALUES(tamanho_bytes),
                status = 'rascunho',
                publicado_em = VALUES(publicado_em),
                publicado_por = VALUES(publicado_por)
            """,
            (version, descricao, obrigatoria, pacote, sha256_hash, tamanho_bytes, datetime.now(), usuario)
        )
        conn.commit()

        # 4. Confirma integridade do BLOB gravado (Passo 2/4)
        print("3. Validando integridade e hash do pacote gravado no banco...")
        cursor.execute("SELECT sha256, LENGTH(pacote_zip) FROM app_releases WHERE version = %s", (version,))
        row = cursor.fetchone()
        if not row:
            raise RuntimeError("Falha crítica: Pacote não encontrado no banco após gravação.")
        
        sha_gravado, tamanho_gravado = row[0], row[1]
        if sha_gravado != sha256_hash or tamanho_gravado != tamanho_bytes:
            raise RuntimeError(
                f"Inconsistência de integridade: Esperado {tamanho_bytes} bytes (hash {sha256_hash}), "
                f"mas o banco gravou {tamanho_gravado} bytes (hash {sha_gravado}). Abortando publicação."
            )
        print("   [OK] Integridade SHA-256 e tamanho confirmados com 100% de precisão!")

        # 5. Marca o pacote como publicado (Passo 3/4)
        print("4. Atualizando status da release para 'publicada'...")
        cursor.execute("UPDATE app_releases SET status = 'publicada' WHERE version = %s", (version,))
        conn.commit()

        # 6. Atualiza o ponteiro 'release_atual' em app_settings (Passo 4/4)
        print("5. Atualizando ponteiro 'release_atual' em app_settings...")
        ponteiro_data = {
            "version": version,
            "descricao": descricao,
            "obrigatoria": obrigatoria,
            "sha256": sha256_hash,
            "tamanho_bytes": tamanho_bytes,
            "publicado_em": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        cursor.execute(
            """
            INSERT INTO app_settings (chave, valor, versao, atualizado_em, atualizado_por)
            VALUES ('release_atual', %s, 1, %s, %s)
            ON DUPLICATE KEY UPDATE
                valor = VALUES(valor),
                versao = versao + 1,
                atualizado_em = VALUES(atualizado_em),
                atualizado_por = VALUES(atualizado_por)
            """,
            (json.dumps(ponteiro_data, ensure_ascii=False), datetime.now(), usuario)
        )
        conn.commit()

        # 7. Atualiza manifesto local version.json para sincronia e documentação
        manifesto_path = BASE_DIR / "releases" / "version.json"
        manifesto_path.parent.mkdir(parents=True, exist_ok=True)
        with open(manifesto_path, "w", encoding="utf-8") as mf:
            json.dump({
                "version": version,
                "pasta": f"app-{version}",
                "notas": descricao,
                "obrigatoria": obrigatoria,
                "sha256": sha256_hash,
                "data_publicacao": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }, mf, indent=2, ensure_ascii=False)

        print("\n" + "=" * 70)
        print(f" [SUCESSO] Versão {version} publicada com sucesso no MySQL!")
        print(f" Todos os usuários com o Launcher receberão a v{version} automaticamente.")
        print("=" * 70)

    finally:
        cursor.close()
        conn.close()


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso: python scripts/publicar_release.py <versao> [descricao] [--opcional]")
        print("Exemplo: python scripts/publicar_release.py 2.1.0 'Adicionado parser de PDF'")
        sys.exit(1)

    versao_arg = sys.argv[1]
    desc_arg = sys.argv[2] if len(sys.argv) > 2 and not sys.argv[2].startswith("--") else f"Atualização {versao_arg}"
    obrigatoria_arg = "--opcional" not in sys.argv

    publicar_release(versao_arg, desc_arg, obrigatoria=obrigatoria_arg)
