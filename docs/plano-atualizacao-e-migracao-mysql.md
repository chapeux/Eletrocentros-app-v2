# Eletrocentros App — Plano de Atualização Automática & Migração dos JSON para MySQL

Documento de arquitetura para resolver dois problemas do modelo atual de distribuição (`.exe` numa pasta de rede):

1. Usuários copiam o app para a Área de Trabalho e ficam presos numa versão antiga.
2. `config.json`, `regras.json`, `seletor.json` e `template_blocks.json` vivem ao lado do `.exe` — cada cópia local vira uma fonte de verdade divergente.

A solução separa duas coisas que hoje estão amarradas ao mesmo lugar: **onde o código roda** e **onde os dados vivem**.

---

## 1. Diagnóstico

| Hoje | Problema |
|---|---|
| `.exe` + `frontend/` inteiros ficam numa pasta de rede | Nada impede o usuário de copiar a pasta toda pro desktop |
| `config.json` / `regras.json` / `seletor.json` / `template_blocks.json` ficam dentro de `frontend/`, ao lado do `.exe` | Uma cópia local = um conjunto de regras próprio, que diverge silenciosamente do resto da empresa |
| `git_sync.py` empurra `regras.json`/`config.json` pro GitHub a cada save | Só funciona se aquela cópia tiver um `.git` válido — uma pasta "só arquivos" falha em silêncio (`except Exception` engolindo o erro) |
| Nenhuma trava de concorrência nos arquivos | Dois usuários salvando regras ao mesmo tempo podem se sobrescrever ou corromper o JSON |

O MySQL (`dcprd036187.weg.net:8502`, banco `bd_eletrocentros_app`) já existe e já registra auditoria em `logs_modificacoes`. Vamos usar essa mesma infraestrutura como fonte de verdade dos dados, e resolver a distribuição do `.exe` com um padrão de launcher + versionamento (o mesmo usado por Steam/Discord, adaptado à escala de uma rede interna).

---

## 2. Visão geral da arquitetura proposta

```
┌─────────────────────────┐        ┌──────────────────────────────┐
│  Atalho na Área de       │        │   Pasta de rede               │
│  Trabalho do usuário     │──────▶│   \\servidor\Eletrocentros\    │
│  (launcher.exe, fixo)    │        │     releases\                 │
└─────────────────────────┘        │       version.json             │
             │                      │       app-2.5.1\ (código)      │
             │ copia/atualiza       └──────────────────────────────┘
             ▼
┌─────────────────────────────┐
│ %LOCALAPPDATA%\Eletrocentros\│
│   app-2.5.0\  (versão antiga)│
│   app-2.5.1\  (versão atual) │◀── executado a partir daqui
│   current.txt                │
└─────────────────────────────┘
             │
             │ lê/grava dados (nunca arquivos locais)
             ▼
┌─────────────────────────────┐
│ MySQL — bd_eletrocentros_app │
│   app_settings (config,      │
│   regras, seletor, templates)│
│   logs_modificacoes (já existe)│
└─────────────────────────────┘
```

Código (versionado, atualiza sozinho) e dados (centralizados, únicos) passam a viajar por caminhos completamente separados. Copiar o `.exe` deixa de ser um risco: ele sempre vai ler/gravar o mesmo banco, não importa de onde rodar.

---

## 3. Parte A — Launcher com atualização automática

### 3.1 Estrutura de pastas na rede

```
\\servidor\Eletrocentros\releases\
    version.json
    app-2.5.0\        (build antiga, mantida como fallback)
    app-2.5.1\         (build atual)
        main.py / app.exe
        frontend\
        assets\
```

`version.json`:

```json
{
  "version": "2.5.1",
  "pasta": "app-2.5.1",
  "notas": "Correção no cálculo de horas ELE + nova aba de Resultados",
  "obrigatoria": true
}
```

`obrigatoria: true` permite marcar updates críticos que bloqueiam o uso da versão antiga (ex.: mudança na regra de cálculo que não pode ficar divergente entre usuários).

### 3.2 O launcher (fica no atalho da Área de Trabalho)

Esse executável é pequeno, muda raramente, e faz só isto:

```python
# launcher.py — compilar separadamente com PyInstaller (launcher.exe)
import json, shutil, subprocess, sys
from pathlib import Path

NETWORK_RELEASES = Path(r"\\servidor\Eletrocentros\releases")
LOCAL_CACHE = Path.home() / "AppData" / "Local" / "Eletrocentros"
CURRENT_FILE = LOCAL_CACHE / "current.txt"


def versao_local() -> str | None:
    if CURRENT_FILE.exists():
        return CURRENT_FILE.read_text(encoding="utf-8").strip()
    return None


def ler_manifesto_rede() -> dict | None:
    try:
        manifest_path = NETWORK_RELEASES / "version.json"
        return json.loads(manifest_path.read_text(encoding="utf-8"))
    except Exception as e:
        print(f"[Launcher] Rede indisponível ou manifesto inválido: {e}")
        return None


def atualizar_se_necessario(manifest: dict) -> Path:
    pasta_alvo = LOCAL_CACHE / manifest["pasta"]

    if pasta_alvo.exists():
        return pasta_alvo  # já está em cache, nada a copiar

    print(f"[Launcher] Baixando nova versão {manifest['version']}...")
    pasta_tmp = LOCAL_CACHE / f"{manifest['pasta']}.tmp"
    if pasta_tmp.exists():
        shutil.rmtree(pasta_tmp)

    shutil.copytree(NETWORK_RELEASES / manifest["pasta"], pasta_tmp)
    pasta_tmp.rename(pasta_alvo)  # troca atômica — só "aparece" pronta

    CURRENT_FILE.parent.mkdir(parents=True, exist_ok=True)
    CURRENT_FILE.write_text(manifest["version"], encoding="utf-8")
    return pasta_alvo


def limpar_versoes_antigas(manter: str, manter_ultimas: int = 2):
    """Mantém só as N últimas versões em cache para não acumular disco."""
    pastas = sorted(
        [p for p in LOCAL_CACHE.iterdir() if p.is_dir() and p.name.startswith("app-")],
        key=lambda p: p.stat().st_mtime, reverse=True,
    )
    for p in pastas[manter_ultimas:]:
        shutil.rmtree(p, ignore_errors=True)


def main():
    manifest = ler_manifesto_rede()

    if manifest is None:
        # Rede fora do ar: roda a última versão em cache, se existir
        v_local = versao_local()
        if v_local is None:
            print("Sem conexão com a rede e nenhuma versão em cache. Contate o suporte.")
            sys.exit(1)
        pasta_alvo = LOCAL_CACHE / f"app-{v_local}"
    else:
        pasta_alvo = atualizar_se_necessario(manifest)
        limpar_versoes_antigas(manifest["version"])

    exe = pasta_alvo / "app.exe"  # ou "main.py" em modo dev
    subprocess.Popen([str(exe)], cwd=str(pasta_alvo))


if __name__ == "__main__":
    main()
```

**Por que isso resolve o problema original:** o usuário nunca mais interage diretamente com o `.exe` do app — só com o launcher, que sempre busca a versão publicada. Copiar o launcher pro desktop é seguro e desejável (é exatamente o "atalho"); copiar a pasta `app-x.y.z` de dentro do cache local não tem efeito, porque os dados não estão mais lá dentro (ver Parte B).

### 3.3 Checklist de publicação de uma nova versão

1. Build gera a nova pasta `app-2.5.2\` dentro de `releases\` (via CI local ou script de build).
2. Atualizar `version.json` por último, só depois que a pasta estiver 100% copiada — isso evita que um usuário pegue o manifesto novo com a pasta ainda incompleta.
3. Não apagar a versão anterior imediatamente — mantenha por alguns dias como rollback manual.

---

## 4. Parte B — Migração dos `.json` para MySQL

### 4.1 Estratégia: JSON como coluna, não normalização completa

`regras.json` tem uma estrutura profundamente aninhada (`area → campos → H/DUR → base/condições/etapas`) usada por um motor de cálculo genérico no `app.js`. Normalizar isso em tabelas relacionais seria um projeto grande e arriscado por si só. A migração pragmática — que resolve o problema real (múltiplas cópias divergentes, sem trava de concorrência) sem reescrever o motor de regras — é guardar cada arquivo inteiro como uma coluna `JSON` no MySQL, com controle de versão otimista:

```sql
CREATE TABLE IF NOT EXISTS app_settings (
    chave           VARCHAR(64)  PRIMARY KEY,
    valor           JSON         NOT NULL,
    versao          INT          NOT NULL DEFAULT 1,
    atualizado_em   DATETIME     NOT NULL,
    atualizado_por  VARCHAR(100) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
```

Cada arquivo atual vira uma linha:

| chave | conteúdo (equivalente a) |
|---|---|
| `config_geral` | `config.json` |
| `regras_calculo` | `regras.json` |
| `seletor_pep_cts` | `seletor.json` |
| `template_blocks` | `template_blocks.json` |

Isso dá, de graça:
- **Uma única fonte de verdade**, não importa quantas cópias do `.exe` existam.
- **Transações reais** — MySQL trata a concorrência, em vez de duas escritas de arquivo pisando uma na outra.
- **Consulta nativa em JSON** se um dia for útil (`JSON_EXTRACT`), sem precisar migrar tudo de uma vez.
- Reaproveita a tabela `logs_modificacoes` que já existe para auditoria — nada muda ali.

### 4.2 Controle de concorrência (lock otimista)

```python
# backend/settings_store.py
import json
from datetime import datetime
from backend.database import get_db_connection, get_current_user


def get_setting(chave: str) -> tuple[dict, int]:
    """Retorna (valor, versao) da configuração. versao é usado para detectar conflitos no save."""
    conn = get_db_connection(with_db=True)
    cursor = conn.cursor()
    cursor.execute("SELECT valor, versao FROM app_settings WHERE chave = %s", (chave,))
    row = cursor.fetchone()
    cursor.close()
    conn.close()
    if row is None:
        return {}, 0
    valor, versao = row
    return json.loads(valor) if isinstance(valor, str) else valor, versao


def save_setting(chave: str, valor: dict, versao_esperada: int, usuario: str | None = None) -> dict:
    """
    Salva com lock otimista: só grava se ninguém alterou a linha desde que este
    usuário carregou os dados (versao_esperada == versao atual no banco).
    """
    usuario = usuario or get_current_user()
    conn = get_db_connection(with_db=True)
    cursor = conn.cursor()

    if versao_esperada == 0:
        # primeira gravação — insere se não existir
        cursor.execute(
            """INSERT INTO app_settings (chave, valor, versao, atualizado_em, atualizado_por)
               VALUES (%s, %s, 1, %s, %s)
               ON DUPLICATE KEY UPDATE chave = chave""",
            (chave, json.dumps(valor, ensure_ascii=False), datetime.now(), usuario),
        )
        conn.commit()
        cursor.close(); conn.close()
        return {"status": "success", "versao": 1}

    cursor.execute(
        """UPDATE app_settings
           SET valor = %s, versao = versao + 1, atualizado_em = %s, atualizado_por = %s
           WHERE chave = %s AND versao = %s""",
        (json.dumps(valor, ensure_ascii=False), datetime.now(), usuario, chave, versao_esperada),
    )
    conn.commit()
    afetadas = cursor.rowcount
    cursor.close()
    conn.close()

    if afetadas == 0:
        return {
            "status": "conflito",
            "message": "Alguém salvou uma versão mais nova enquanto você editava. Recarregue e refaça a alteração.",
        }
    return {"status": "success", "versao": versao_esperada + 1}
```

No `app.js`, o único ajuste de UX necessário é tratar o retorno `"conflito"` (hoje inexistente, porque escrever em arquivo local nunca "falha" desse jeito) — mostrar um aviso e recarregar os dados antes de deixar salvar de novo.

### 4.3 Ajustando o `AppAPI` — mesma assinatura, fonte diferente

O ponto forte dessa migração: o frontend (`app.js`) **não precisa mudar** — `get_config`, `save_config`, `get_regras`, `save_regras`, `get_seletor`, `save_seletor`, `get_template_blocks`, `save_template_blocks` continuam existindo com a mesma assinatura, só trocam o "por baixo dos panos":

```python
# main.py — trechos alterados
from backend.settings_store import get_setting, save_setting

class AppAPI:
    def get_config(self) -> dict:
        valor, versao = get_setting("config_geral")
        self._versao_config = versao  # guardado em memória p/ o próximo save
        return valor

    def save_config(self, config_data: dict) -> dict:
        resultado = save_setting("config_geral", config_data, getattr(self, "_versao_config", 0))
        if resultado["status"] == "success":
            sync_github_async(resumo="Atualização de config (MySQL)")  # backup opcional, ver seção 4.5
        return resultado
```

O mesmo padrão se repete para `regras_calculo`, `seletor_pep_cts` e `template_blocks`.

### 4.4 Script de migração (roda uma vez)

```python
# scripts/migrar_json_para_mysql.py
import json
from pathlib import Path
from datetime import datetime
from backend.database import get_db_connection, init_db

FRONTEND_DIR = Path(__file__).resolve().parent.parent / "frontend"

ARQUIVOS = {
    "config_geral":     FRONTEND_DIR / "config.json",
    "regras_calculo":   FRONTEND_DIR / "regras.json",
    "seletor_pep_cts":  FRONTEND_DIR / "seletor.json",
    "template_blocks":  FRONTEND_DIR / "template_blocks.json",
}

def main():
    init_db()  # garante banco + tabela logs já existentes
    conn = get_db_connection(with_db=True)
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS app_settings (
            chave VARCHAR(64) PRIMARY KEY,
            valor JSON NOT NULL,
            versao INT NOT NULL DEFAULT 1,
            atualizado_em DATETIME NOT NULL,
            atualizado_por VARCHAR(100) NOT NULL
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
    """)

    for chave, caminho in ARQUIVOS.items():
        if not caminho.exists():
            print(f"[Migração] {caminho} não encontrado, pulando.")
            continue
        dados = json.loads(caminho.read_text(encoding="utf-8"))
        cursor.execute(
            """INSERT INTO app_settings (chave, valor, versao, atualizado_em, atualizado_por)
               VALUES (%s, %s, 1, %s, 'migracao_inicial')
               ON DUPLICATE KEY UPDATE valor = VALUES(valor)""",
            (chave, json.dumps(dados, ensure_ascii=False), datetime.now()),
        )
        print(f"[Migração] {chave} carregado a partir de {caminho.name} ({len(json.dumps(dados))} bytes).")

    conn.commit()
    cursor.close()
    conn.close()
    print("[Migração] Concluída.")

if __name__ == "__main__":
    main()
```

Rodar uma única vez, a partir dos arquivos que já estão na pasta de rede oficial (nunca de uma cópia local).

### 4.5 Remoção do `git_sync.py`

O `git_sync.py` foi criado só para manter o repositório Git atualizado enquanto os `.json` eram a única fonte de verdade — era a forma de ter algum histórico/backup do que estava sendo salvo. Com os dados centralizados em `app_settings` (fonte única, sem risco de divergência entre cópias) e o histórico de alterações já coberto por `logs_modificacoes`, ele passa a ser redundante.

Remover ao final da migração (depois que o `AppAPI` já estiver lendo/gravando 100% do MySQL, e não mais dos arquivos locais):

- Apagar `backend/git_sync.py`.
- Remover o `import` e as chamadas `sync_github_async(...)` de `main.py` (em `save_config`, `save_regras`, `save_seletor`, `save_template_blocks`).
- Se o repositório Git da pasta de rede não tiver mais nenhuma outra finalidade, ele também pode ser descontinuado — mas isso é opcional e não bloqueia o restante do plano.

### 4.6 Resiliência a rede/MySQL fora do ar

Como os dados passam a depender do MySQL estar acessível, vale um fallback explícito: se `get_setting` falhar (timeout de conexão), o `AppAPI` cai para o último `frontend/*.json` conhecido (o cache de leitura mencionado acima) e **desabilita a gravação** (bloqueia o botão salvar com aviso "Sem conexão com o banco — alterações não podem ser salvas agora"), em vez de deixar o usuário editar algo que vai se perder.

---

## 5. Ordem de execução recomendada

1. **Lock de escrita nos arquivos atuais** (curto prazo, sem MySQL ainda): trocar `open(...).write()` direto por escrita atômica (`.tmp` + `os.replace`) — reduz risco de corrupção enquanto o resto não está pronto.
2. **Publicar o launcher** (Parte A) e passar a distribuir só ele nos atalhos — resolve o problema de versão desatualizada primeiro, é independente da migração de dados.
3. **Criar `app_settings`, rodar o script de migração, ajustar `AppAPI`** (Parte B) — pode ser feito com o app ainda rodando da rede (sem o launcher), validando em paralelo antes de depender 100% do banco.
4. **Trocar a fonte de leitura do frontend para o MySQL em produção**, mantendo os `.json` locais só como cache de leitura/fallback (ver §4.6).
5. **Descomissionar a leitura direta de `frontend/*.json` como fonte de escrita e remover o `git_sync.py`** (§4.5) — a partir daqui, uma cópia isolada do app não tem mais como divergir, porque não há mais "dados próprios" para copiar nem sincronizar.

---

## 6. Riscos e mitigação

| Risco | Mitigação |
|---|---|
| MySQL indisponível bloqueia todo mundo | Fallback de leitura em cache local (`.json`) + bloqueio de escrita, não de leitura |
| Credenciais do MySQL hardcoded em `backend/database.py` | Fora do escopo deste documento, mas vale mover para variável de ambiente/arquivo de configuração fora do controle de versão numa próxima iteração |
| Migração perder dados em produção | Rodar o script de migração primeiro contra uma cópia/banco de teste; validar contagem de chaves e um diff campo a campo antes de apontar o `AppAPI` para o banco real |
| Conflito de edição otimista confundir o usuário | Mensagem clara + recarregar automaticamente os dados mais recentes ao detectar conflito, em vez de só travar |
