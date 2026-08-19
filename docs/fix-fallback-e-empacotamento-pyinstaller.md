# Eletrocentros App — Correção do Fallback HTTP & Empacotamento com PyInstaller

Cobre dois ajustes ligados ao erro reportado (`WinError 10048` na porta 8000, precedido por `pywebview não encontrado`):

1. **Correção imediata de código** — o fallback HTTP não deveria travar quando já existe uma instância rodando, e o loop principal não deveria consumir CPU à toa.
2. **Correção da causa raiz** — parar de depender do Python/pywebview instalado em cada máquina, empacotando o app com PyInstaller.

---

## 1. Diagnóstico confirmado

```
[GUI Python] pywebview não encontrado no ambiente Python.       ← sintoma: dependência não instalada nessa máquina
[GUI Python] Iniciando servidor HTTP local de fallback...
OSError: [WinError 10048] ... apenas uma utilização de cada endereço de soquete ...  ← erro real: porta 8000 já ocupada
```

O `main.py` hoje roda através do Python instalado no sistema de cada usuário (`C:\Program Files\Python312\...`), não como um executável autocontido. Isso explica por que "só no seu PC funciona": o pacote `pywebview` está instalado no seu Python local, mas não no das outras máquinas.

Quando `pywebview` falta, o app cai no fallback (servidor HTTP + navegador). Esse fallback tem um bug: o processo nunca termina sozinho —

```python
try:
    while True:
        pass                      # busy-wait: 100% de um núcleo de CPU, para sempre
except KeyboardInterrupt:
    ...
```

Fechar a aba do navegador não encerra o processo Python. Se o app for aberto de novo (segundo clique, launcher relançando), a instância anterior ainda está viva segurando a porta 8000 — daí o `WinError 10048`.

---

## 2. Correção imediata no `main.py`

### 2.1 Detectar instância já rodando, em vez de travar

```python
import socket
import time

def porta_em_uso(port: int, host: str = "127.0.0.1") -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.settimeout(0.3)
        return s.connect_ex((host, port)) == 0


class ReusableTCPServer(socketserver.TCPServer):
    allow_reuse_address = True   # evita "endereço em uso" logo após um fechamento malfeito


def run_local_server(port: int = 8000):
    os.chdir(FRONTEND_DIR)
    handler = AppHTTPRequestHandler
    with ReusableTCPServer(("", port), handler) as httpd:
        print(f"[Python HTTP Server] Servindo frontend e API em http://localhost:{port}")
        httpd.serve_forever()
```

E no `launch_app()`, dentro do `except ImportError:`:

```python
except ImportError:
    print("[GUI Python] pywebview não encontrado no ambiente Python.")
    print("[GUI Python] Iniciando servidor HTTP local de fallback...")

    port = 8000
    url = f"http://localhost:{port}/index.html"

    if porta_em_uso(port):
        print("[GUI Python] Já existe uma instância rodando nesta máquina — abrindo nela.")
    else:
        server_thread = threading.Thread(target=run_local_server, args=(port,), daemon=True)
        server_thread.start()
        time.sleep(0.3)  # dá tempo do bind acontecer antes de abrir o navegador

    print(f"[GUI Python] Abrindo navegador padrão em {url}...")
    webbrowser.open(url)

    try:
        while True:
            time.sleep(1)   # substitui o `pass` — não consome CPU à toa
    except KeyboardInterrupt:
        print("\n[GUI Python] Aplicação encerrada pelo usuário.")
```

Isso resolve o crash reportado. Mas é um remendo em cima do sintoma — o motivo de cair no fallback continua existindo. A correção de verdade é a seção 3.

---

## 3. Empacotamento com PyInstaller (elimina a dependência do Python do sistema)

### 3.1 Por que `--onedir`, não `--onefile`

| | `--onefile` | `--onedir` (recomendado aqui) |
|---|---|---|
| Distribuição | 1 `.exe` só | Pasta com `.exe` + arquivos ao lado |
| Inicialização | Mais lenta (extrai tudo em `%TEMP%` a cada execução) | Instantânea |
| Antivírus corporativo | Mais falsos positivos (padrão comum de malware) | Menos alertas |
| Encaixe com o pipeline de releases via MySQL | Teria que re-extrair o `.exe` de dentro do zip antes de rodar | O `.zip` publicado em `app_releases` já é literalmente essa pasta — zero passo extra |

Como o launcher já baixa e extrai um `.zip` do MySQL, `--onedir` encaixa diretamente: a pasta de saída do PyInstaller **é** o pacote que vai para `app_releases.pacote_zip`.

### 3.2 Ajuste necessário no `main.py` antes de empacotar

`BASE_DIR` hoje é `Path(__file__).resolve().parent`, o que quebra dentro de um build congelado (o PyInstaller reorganiza os caminhos). Trocar por:

```python
if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).resolve().parent
else:
    BASE_DIR = Path(__file__).resolve().parent

FRONTEND_DIR = BASE_DIR / "frontend"
```

Isso garante que, rodando como `app.exe`, `FRONTEND_DIR`/`ICON_PATH`/`CONFIG_FILE` continuem apontando para as pastas ao lado do executável (que o PyInstaller copia via `--add-data`, ver 3.4), e não para dentro de um pacote temporário.

### 3.3 Dependências que precisam de `--hidden-import`

O PyInstaller não detecta sozinho os plugins que o `pywebview` carrega dinamicamente no Windows (ele escolhe o motor Edge WebView2/CEF em tempo de execução), nem alguns submódulos do `openpyxl` usados só sob certas condições:

```
--hidden-import=webview.platforms.winforms
--hidden-import=webview.platforms.edgechromium
--hidden-import=clr_loader
--hidden-import=mysql.connector.locales.eng
--hidden-import=openpyxl.cell._writer
```

### 3.4 Comando de build

```powershell
pyinstaller ^
  --name app ^
  --onedir ^
  --windowed ^
  --icon assets\icone.ico ^
  --add-data "frontend;frontend" ^
  --add-data "assets;assets" ^
  --hidden-import webview.platforms.winforms ^
  --hidden-import webview.platforms.edgechromium ^
  --hidden-import clr_loader ^
  --hidden-import mysql.connector.locales.eng ^
  --hidden-import openpyxl.cell._writer ^
  main.py
```

Saída: `dist\app\` — contendo `app.exe`, `frontend\`, `assets\` e todas as DLLs/dependências. Essa é a pasta inteira que deve ser zipada e publicada (ver §4).

### 3.5 `app.spec` (alternativa reutilizável ao comando acima)

Depois do primeiro build, o PyInstaller já gera um `app.spec` editável — mais prático para builds repetidos que o comando longo:

```python
# app.spec
# -*- mode: python ; coding: utf-8 -*-

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('frontend', 'frontend'),
        ('assets', 'assets'),
    ],
    hiddenimports=[
        'webview.platforms.winforms',
        'webview.platforms.edgechromium',
        'clr_loader',
        'mysql.connector.locales.eng',
        'openpyxl.cell._writer',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='app',
    debug=False,
    strip=False,
    upx=True,
    console=False,       # sem console — janela desktop limpa
    icon='assets\\icone.ico',
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    name='app',
)
```

Build subsequente vira só:

```powershell
pyinstaller app.spec
```

### 3.6 Checklist de validação (obrigatório antes de publicar)

Testar em uma máquina **sem Python instalado** — é o único jeito de confirmar que a dependência de ambiente foi realmente eliminada:

1. Copiar só `dist\app\` para essa máquina (sem levar nenhum Python junto).
2. Rodar `app.exe` e confirmar no log que a janela nativa abre (`[GUI Python] Inicializando janela nativa com PyWebView...`) — **sem** cair no fallback.
3. Testar salvar uma regra (grava no MySQL) e exportar Excel (valida `openpyxl` embutido).
4. Fechar e abrir de novo para garantir que não sobra processo travando porta nenhuma (não deveria nem chegar a usar o fallback, mas vale confirmar que o processo antigo encerra ao fechar a janela nativa).

---

## 4. Encaixe com o pipeline de releases via MySQL

Nenhuma mudança no schema (`app_releases`) ou no launcher é necessária — só o conteúdo do `.zip` publicado muda: em vez de zipar a pasta com `main.py` solto, zipa-se `dist\app\` (o resultado do PyInstaller).

```python
# scripts/publicar_release.py — único ajuste na constante do topo
BUILD_DIR = Path("dist/app")   # antes apontava para uma pasta com main.py + dependências soltas
```

E o launcher, que já espera um `app.exe` dentro do pacote extraído (`pasta_alvo / "app.exe"`), não precisa de nenhum ajuste — já está pronto para essa mudança.

---

## 5. Riscos e notas

| Risco | Nota |
|---|---|
| Antivírus corporativo sinalizando o `.exe` | Mais raro com `--onedir` que com `--onefile`, mas ainda pode acontecer na primeira execução em cada máquina (heurística "executável desconhecido"). Se for recorrente, considerar assinatura de código (certificado da empresa) |
| Tamanho do pacote aumenta bastante (Python + libs embutidos) | Esperado — normalmente vai de poucos KB (só código-fonte) para dezenas de MB. Reforça o cuidado com `max_allowed_packet` do MySQL já anotado no documento de releases |
| `webview2` (motor do Edge Chromium) precisa estar presente no Windows do usuário | Vem pré-instalado por padrão no Windows 10/11 atualizados; só é um problema em máquinas muito desatualizadas — nesse caso o app cai no fallback normalmente, então não é um travamento, é só uma degradação |
| Build precisa ser feito numa máquina Windows (PyInstaller não faz cross-compile de Windows a partir de Linux/Mac) | Gerar o build sempre na mesma máquina/ambiente de referência, para evitar builds inconsistentes |
