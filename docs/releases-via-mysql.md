# Eletrocentros App — Distribuição de Releases via MySQL (sem pasta de rede)

Substitui a lógica de atualização baseada em `\\servidor\Eletrocentros\releases\` por um fluxo 100% via MySQL. Motivação: parte dos usuários não tem acesso à pasta de rede, mas todos já precisam de conectividade com o banco (`dcprd036187.weg.net:8502` / `bd_eletrocentros_app`) para o app funcionar normalmente.

Pré-requisito: a tabela `app_settings` (chave/valor JSON com versionamento) já implementada na migração dos `.json` de configuração.

---

## 1. O que muda

| Antes (pasta de rede) | Agora (MySQL) |
|---|---|
| `version.json` numa pasta compartilhada | Linha `release_atual` em `app_settings` |
| Pastas `app-x.y.z\` com os arquivos do build | Pacotes `.zip` guardados como `LONGBLOB` em `app_releases` |
| Launcher faz `shutil.copytree` de um caminho `\\servidor\...` | Launcher faz `SELECT` do BLOB e extrai localmente |
| Sem verificação de integridade do que foi copiado | SHA-256 conferido antes de aplicar |

O único ponto que continua "físico" é a entrega inicial do próprio `launcher.exe` na primeira instalação (download único via intranet, e-mail ou GPO) — depois disso, tudo passa a fluir só pelo banco.

---

## 2. Schema

```sql
CREATE TABLE IF NOT EXISTS app_releases (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    version         VARCHAR(20)  NOT NULL UNIQUE,
    descricao       TEXT,
    obrigatoria     BOOLEAN      NOT NULL DEFAULT FALSE,
    pacote_zip      LONGBLOB     NOT NULL,   -- build inteiro (app + frontend) compactado
    sha256          CHAR(64)     NOT NULL,   -- integridade do download
    tamanho_bytes   BIGINT       NOT NULL,
    status          ENUM('rascunho','publicada','revogada') NOT NULL DEFAULT 'rascunho',
    publicado_em    DATETIME     NOT NULL,
    publicado_por   VARCHAR(100) NOT NULL,
    INDEX idx_status (status)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
```

`app_settings` recebe um ponteiro leve, para o launcher não precisar escanear BLOBs a cada abertura:

```json
// chave = "release_atual"
{ "version": "2.5.2", "obrigatoria": true }
```

`LONGBLOB` suporta até 4GB — muito acima do necessário para este app; ainda assim, ver §6 sobre limites práticos de `max_allowed_packet` no MySQL.

---

## 3. Publicando uma nova versão

Script rodado manualmente (ou via CI) a cada release, a partir da pasta de build local do desenvolvedor:

```python
# scripts/publicar_release.py
import hashlib
import json
import zipfile
from datetime import datetime
from io import BytesIO
from pathlib import Path

from backend.database import get_db_connection, get_current_user

BUILD_DIR = Path("dist/app-2.5.2")  # pasta com o build já compilado (exe + frontend)
VERSION = "2.5.2"
DESCRICAO = "Correção no cálculo de horas ELE + nova aba de Resultados"
OBRIGATORIA = True


def zipar_build(pasta: Path) -> bytes:
    buffer = BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for arquivo in pasta.rglob("*"):
            if arquivo.is_file():
                zf.write(arquivo, arquivo.relative_to(pasta))
    return buffer.getvalue()


def main():
    pacote = zipar_build(BUILD_DIR)
    sha256 = hashlib.sha256(pacote).hexdigest()
    usuario = get_current_user()

    conn = get_db_connection(with_db=True)
    cursor = conn.cursor()

    # 1. Grava o pacote como rascunho primeiro
    cursor.execute(
        """INSERT INTO app_releases
           (version, descricao, obrigatoria, pacote_zip, sha256, tamanho_bytes, status, publicado_em, publicado_por)
           VALUES (%s, %s, %s, %s, %s, %s, 'rascunho', %s, %s)""",
        (VERSION, DESCRICAO, OBRIGATORIA, pacote, sha256, len(pacote), datetime.now(), usuario),
    )
    conn.commit()

    # 2. Confirma que o BLOB gravou íntegro antes de publicar de vez
    cursor.execute("SELECT sha256, LENGTH(pacote_zip) FROM app_releases WHERE version = %s", (VERSION,))
    sha_gravado, tamanho_gravado = cursor.fetchone()
    assert sha_gravado == sha256 and tamanho_gravado == len(pacote), "Pacote gravado difere do original — abortando."

    # 3. Só agora marca como publicada
    cursor.execute("UPDATE app_releases SET status = 'publicada' WHERE version = %s", (VERSION,))

    # 4. Atualiza o ponteiro por último — é o que o launcher enxerga
    cursor.execute(
        """INSERT INTO app_settings (chave, valor, versao, atualizado_em, atualizado_por)
           VALUES ('release_atual', %s, 1, %s, %s)
           ON DUPLICATE KEY UPDATE valor = VALUES(valor), versao = versao + 1,
                                    atualizado_em = VALUES(atualizado_em), atualizado_por = VALUES(atualizado_por)""",
        (json.dumps({"version": VERSION, "obrigatoria": OBRIGATORIA}, ensure_ascii=False), datetime.now(), usuario),
    )
    conn.commit()
    cursor.close()
    conn.close()
    print(f"[Publicação] Versão {VERSION} publicada com sucesso ({len(pacote)/1024:.1f} KB).")


if __name__ == "__main__":
    main()
```

A ordem dos passos 1→4 importa: o ponteiro em `app_settings` só muda **depois** que o BLOB já está confirmado e gravado com `status='publicada'`. Isso evita que um usuário pegue o ponteiro apontando para uma versão cujo pacote ainda não terminou de gravar.

---

## 4. O launcher

```python
# launcher.py — compilado separadamente com PyInstaller (launcher.exe)
import hashlib
import json
import subprocess
import sys
import zipfile
from io import BytesIO
from pathlib import Path

from backend.database import get_db_connection

LOCAL_CACHE = Path.home() / "AppData" / "Local" / "Eletrocentros"
CURRENT_FILE = LOCAL_CACHE / "current.txt"


def versao_local() -> str | None:
    if CURRENT_FILE.exists():
        return CURRENT_FILE.read_text(encoding="utf-8").strip()
    return None


def buscar_ponteiro_release() -> dict | None:
    try:
        conn = get_db_connection(with_db=True)
        cursor = conn.cursor()
        cursor.execute("SELECT valor FROM app_settings WHERE chave = 'release_atual'")
        row = cursor.fetchone()
        cursor.close()
        conn.close()
        if row is None:
            return None
        return json.loads(row[0]) if isinstance(row[0], str) else row[0]
    except Exception as e:
        print(f"[Launcher] Não foi possível consultar o MySQL: {e}")
        return None


def baixar_pacote(version: str) -> tuple[bytes, str]:
    conn = get_db_connection(with_db=True)
    cursor = conn.cursor()
    cursor.execute(
        "SELECT pacote_zip, sha256 FROM app_releases WHERE version = %s AND status = 'publicada'",
        (version,),
    )
    row = cursor.fetchone()
    cursor.close()
    conn.close()
    if row is None:
        raise RuntimeError(f"Versão {version} não encontrada ou não publicada no banco.")
    return row[0], row[1]


def aplicar_pacote(pacote_bytes: bytes, version: str) -> Path:
    pasta_alvo = LOCAL_CACHE / f"app-{version}"
    pasta_tmp = LOCAL_CACHE / f"app-{version}.tmp"

    if pasta_tmp.exists():
        import shutil
        shutil.rmtree(pasta_tmp)
    pasta_tmp.mkdir(parents=True)

    with zipfile.ZipFile(BytesIO(pacote_bytes)) as zf:
        zf.extractall(pasta_tmp)

    pasta_tmp.rename(pasta_alvo)  # troca atômica
    CURRENT_FILE.parent.mkdir(parents=True, exist_ok=True)
    CURRENT_FILE.write_text(version, encoding="utf-8")
    return pasta_alvo


def limpar_versoes_antigas(manter_ultimas: int = 2):
    pastas = sorted(
        [p for p in LOCAL_CACHE.iterdir() if p.is_dir() and p.name.startswith("app-")],
        key=lambda p: p.stat().st_mtime, reverse=True,
    )
    import shutil
    for p in pastas[manter_ultimas:]:
        shutil.rmtree(p, ignore_errors=True)


def main():
    ponteiro = buscar_ponteiro_release()

    if ponteiro is None:
        # Sem conexão com o MySQL: roda a última versão em cache, se existir
        v_local = versao_local()
        if v_local is None:
            print("Sem conexão com o servidor e nenhuma versão instalada. Contate o suporte de TI.")
            sys.exit(1)
        pasta_alvo = LOCAL_CACHE / f"app-{v_local}"
    else:
        v_local = versao_local()
        if v_local == ponteiro["version"]:
            pasta_alvo = LOCAL_CACHE / f"app-{v_local}"
        else:
            print(f"[Launcher] Baixando versão {ponteiro['version']}...")
            pacote_bytes, sha_esperado = baixar_pacote(ponteiro["version"])

            if hashlib.sha256(pacote_bytes).hexdigest() != sha_esperado:
                print("[Launcher] Pacote corrompido no download. Tentando novamente...")
                pacote_bytes, sha_esperado = baixar_pacote(ponteiro["version"])
                if hashlib.sha256(pacote_bytes).hexdigest() != sha_esperado:
                    raise RuntimeError("Pacote consistentemente corrompido — abortando atualização.")

            pasta_alvo = aplicar_pacote(pacote_bytes, ponteiro["version"])
            limpar_versoes_antigas()

    exe = pasta_alvo / "app.exe"
    subprocess.Popen([str(exe)], cwd=str(pasta_alvo))


if __name__ == "__main__":
    main()
```

Diferenças-chave em relação à versão baseada em pasta de rede:
- `ler_manifesto_rede()` vira `buscar_ponteiro_release()` — troca leitura de arquivo por `SELECT`.
- `shutil.copytree(...)` vira `baixar_pacote()` + `zf.extractall(...)` — o "arquivo" agora é um BLOB, extraído localmente.
- Verificação de SHA-256 antes de aplicar, com uma tentativa extra de download em caso de corrupção — algo que não existia no fluxo por pasta de rede (cópia de arquivo local é praticamente sempre íntegra; download de BLOB via rede corporativa merece essa checagem).

---

## 5. Checklist de corte (rede → MySQL)

1. Criar a tabela `app_releases` no banco (script no §2).
2. Publicar a versão atual do app através do `publicar_release.py`, usando o build que hoje está na pasta de rede.
3. Trocar o `launcher.exe` atual pela versão que consulta o MySQL (§4) e recompilar.
4. Redistribuir o novo `launcher.exe` nos atalhos dos usuários (mesmo processo já usado para entregar o launcher original).
5. Validar com pelo menos um usuário que **não tem acesso à pasta de rede** — esse é o teste real do que motivou a mudança.
6. Só depois de confirmado, desligar a publicação via pasta de rede (pode manter os arquivos lá por um tempo como histórico, sem uso ativo).

---

## 6. Riscos e limites práticos

| Risco | Mitigação |
|---|---|
| `max_allowed_packet` do MySQL bloqueia BLOBs grandes (padrão costuma ser 4–64MB) | Confirmar/configurar esse parâmetro no servidor para acomodar o tamanho do build zipado; validar antes de publicar a primeira release |
| Banco cresce a cada versão publicada | Manter só as últimas 2–3 em `status='publicada'`; mover releases antigas para `status='revogada'` (mantém histórico sem carregar o BLOB no fluxo normal) em vez de `DELETE` |
| `mysql-connector-python` embutido aumenta o tamanho do `launcher.exe` | Aceitável — o launcher é compilado raramente, diferente do app principal |
| Download do BLOB falha no meio (rede instável) | A checagem de SHA-256 já cobre isso — se não bater, tenta de novo antes de aplicar |
| Servidor MySQL fora do ar impede qualquer atualização | Fallback já embutido no launcher: roda a última versão em cache local se não conseguir consultar o banco |
