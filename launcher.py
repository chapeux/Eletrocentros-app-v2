"""
Eletrocentros App — Launcher & Auto-Updater Inteligente
Verifica novas versões publicadas na pasta de rede da WEG, sincroniza
com o cache local (%LOCALAPPDATA%\\EletrocentrosApp) e executa o aplicativo.
"""

import os
import sys
import json
import shutil
import subprocess
import time
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox

# 1. Definição correta do diretório base (funciona tanto como script quanto compilado .exe)
if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).resolve().parent
    INTERNAL_DIR = Path(getattr(sys, "_MEIPASS", BASE_DIR))
else:
    BASE_DIR = Path(__file__).resolve().parent
    INTERNAL_DIR = BASE_DIR

# Caminho principal na rede da WEG (com fallback para pasta local de releases)
NETWORK_DIR_PRIMARY = Path(
    r"Q:\GROUPS\BR_SC_ITJ_WAU_DPTO_PRODUCAO\Processos WAU Chaves Especiais e Acionamentos\10 - PASTAS PESSOAIS\Luan Schappo\Salas Elétricas\Eletrocentros-app-v2\releases"
)
NETWORK_DIR_FALLBACK = BASE_DIR / "releases"

LOCAL_APPDATA = Path(os.environ.get("LOCALAPPDATA", Path.home() / "AppData" / "Local"))
LOCAL_CACHE_DIR = LOCAL_APPDATA / "EletrocentrosApp"
CURRENT_VERSION_FILE = LOCAL_CACHE_DIR / "current_version.txt"
ICON_PATH = INTERNAL_DIR / "assets" / "icone.ico"
if not ICON_PATH.exists():
    ICON_PATH = BASE_DIR / "assets" / "icone.ico"


def obter_pasta_releases_rede() -> Path:
    """Retorna o diretório de releases acessível (rede ou local)."""
    try:
        if NETWORK_DIR_PRIMARY.exists():
            return NETWORK_DIR_PRIMARY
    except Exception:
        pass
    return NETWORK_DIR_FALLBACK


def ler_manifesto_rede(releases_dir: Path) -> dict | None:
    """Lê o arquivo version.json do repositório de releases."""
    manifest_file = releases_dir / "version.json"
    try:
        if manifest_file.exists():
            with open(manifest_file, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception as e:
        print(f"[Launcher] Aviso ao ler manifesto de rede ({manifest_file}): {e}")
    return None


def ler_versao_local() -> str | None:
    """Retorna a versão atualmente instalada no cache local."""
    try:
        if CURRENT_VERSION_FILE.exists():
            return CURRENT_VERSION_FILE.read_text(encoding="utf-8").strip()
    except Exception:
        pass
    return None


def limpar_versoes_antigas(manter_ultimas: int = 2):
    """Remove versões obsoletas do cache local, mantendo apenas as últimas para economia de disco."""
    try:
        if not LOCAL_CACHE_DIR.exists():
            return
        pastas = sorted(
            [p for p in LOCAL_CACHE_DIR.iterdir() if p.is_dir() and p.name.startswith("app-")],
            key=lambda p: p.stat().st_mtime,
            reverse=True,
        )
        for pasta_obsoleta in pastas[manter_ultimas:]:
            try:
                shutil.rmtree(pasta_obsoleta, ignore_errors=True)
                print(f"[Launcher] Versão antiga limpa do cache: {pasta_obsoleta.name}")
            except Exception:
                pass
    except Exception as e:
        print(f"[Launcher] Aviso ao limpar versões antigas: {e}")


class UpdateSplash:
    """Janela moderna de progresso de atualização visual."""

    def __init__(self, versao_nova: str):
        self.root = tk.Tk()
        self.root.title("Eletrocentros App — Atualização")
        self.root.geometry("450x170")
        self.root.resizable(False, False)
        self.root.configure(bg="#0e1726")

        # Centralizar na tela
        self.root.update_idletasks()
        w = self.root.winfo_width()
        h = self.root.winfo_height()
        ws = self.root.winfo_screenwidth()
        hs = self.root.winfo_screenheight()
        x = (ws // 2) - (w // 2)
        y = (hs // 2) - (h // 2)
        self.root.geometry(f"+{x}+{y}")

        # Ícone da janela
        if ICON_PATH.exists():
            try:
                self.root.iconbitmap(str(ICON_PATH))
            except Exception:
                pass

        # Estilo escuro moderno
        style = ttk.Style(self.root)
        style.theme_use("clam")
        style.configure(
            "Update.Horizontal.TProgressbar",
            troughcolor="#1e293b",
            background="#0075c9",
            lightcolor="#0075c9",
            darkcolor="#0075c9",
            bordercolor="#1e293b"
        )

        title_lbl = tk.Label(
            self.root,
            text="Atualizando Eletrocentros App",
            font=("Segoe UI", 12, "bold"),
            fg="#f8fafc",
            bg="#0e1726"
        )
        title_lbl.pack(pady=(16, 4))

        self.info_lbl = tk.Label(
            self.root,
            text=f"Instalando a nova versão {versao_nova} a partir da rede WEG...",
            font=("Segoe UI", 9),
            fg="#94a3b8",
            bg="#0e1726"
        )
        self.info_lbl.pack(pady=(0, 14))

        self.progress = ttk.Progressbar(
            self.root,
            style="Update.Horizontal.TProgressbar",
            mode="indeterminate",
            length=380
        )
        self.progress.pack(pady=(0, 12))
        self.progress.start(15)

        self.status_lbl = tk.Label(
            self.root,
            text="Copiando arquivos e sincronizando componentes...",
            font=("Segoe UI", 8, "italic"),
            fg="#64748b",
            bg="#0e1726"
        )
        self.status_lbl.pack()

        self.root.update()

    def set_status(self, text: str):
        try:
            self.status_lbl.config(text=text)
            self.root.update()
        except Exception:
            pass

    def close(self):
        try:
            self.progress.stop()
            self.root.destroy()
        except Exception:
            pass


def copiar_arvore_com_progresso(origem: Path, destino: Path, splash: UpdateSplash | None = None):
    """Copia a árvore de arquivos de forma atômica e resiliente, ignorando arquivos temporários ou de build."""
    destino.mkdir(parents=True, exist_ok=True)
    itens_ignorados = {".git", "__pycache__", "build", "dist", ".tmp"}
    
    for item in origem.iterdir():
        if item.name in itens_ignorados or item.name.endswith(".tmp") or item.name.endswith(".spec"):
            continue
        dest_item = destino / item.name
        if item.is_dir():
            if splash:
                splash.set_status(f"Copiando pasta: {item.name}...")
            shutil.copytree(item, dest_item, dirs_exist_ok=True, ignore=shutil.ignore_patterns("__pycache__", "*.pyc", "*.tmp"))
        else:
            if splash:
                splash.set_status(f"Copiando arquivo: {item.name}...")
            shutil.copy2(item, dest_item)


def atualizar_aplicativo(releases_dir: Path, manifest: dict) -> Path:
    """Executa a atualização atômica para a nova versão especificada no manifesto."""
    versao = manifest.get("version", "2.0.0")
    nome_pasta = manifest.get("pasta", f"app-{versao}")
    pasta_origem = releases_dir / nome_pasta
    pasta_destino = LOCAL_CACHE_DIR / nome_pasta

    # Se a pasta de release não existir na pasta releases, usa a pasta raiz da rede como origem
    if not pasta_origem.exists():
        pasta_origem = NETWORK_DIR_PRIMARY.parent if NETWORK_DIR_PRIMARY.parent.exists() else BASE_DIR

    splash = UpdateSplash(versao)

    pasta_tmp = LOCAL_CACHE_DIR / f"{nome_pasta}.tmp"
    if pasta_tmp.exists():
        shutil.rmtree(pasta_tmp, ignore_errors=True)

    try:
        splash.set_status("Transferindo nova versão...")
        copiar_arvore_com_progresso(pasta_origem, pasta_tmp, splash)

        # Troca atômica (remove destino antigo se houver e renomeia .tmp)
        if pasta_destino.exists():
            shutil.rmtree(pasta_destino, ignore_errors=True)
        
        pasta_tmp.rename(pasta_destino)

        # Salva o número da versão atual
        LOCAL_CACHE_DIR.mkdir(parents=True, exist_ok=True)
        CURRENT_VERSION_FILE.write_text(versao, encoding="utf-8")
        time.sleep(0.3)
        return pasta_destino

    finally:
        splash.close()


def obter_python_exe() -> str:
    """Retorna o caminho do interpretador Python para executar main.py em modo dev."""
    if not getattr(sys, "frozen", False):
        return sys.executable

    import shutil as sh
    py = sh.which("pythonw") or sh.which("python")
    if py:
        return py

    candidatos = [
        Path(os.environ.get("LOCALAPPDATA", "")) / "Programs" / "Python" / "Python314" / "pythonw.exe",
        Path(os.environ.get("LOCALAPPDATA", "")) / "Programs" / "Python" / "Python314" / "python.exe",
        Path(r"C:\Program Files\Python314\pythonw.exe"),
        Path(r"C:\Program Files\Python314\python.exe"),
        Path(r"C:\Python314\pythonw.exe"),
    ]
    for c in candidatos:
        if c.exists():
            return str(c)
    return "python"


def encontrar_executavel(pasta_app: Path) -> list[str]:
    """Descobre o comando correto para iniciar o executável ou script Python."""
    # 1. Executável compilado
    exe_candidatos = [
        pasta_app / "EletrocentrosApp.exe",
        pasta_app / "app.exe",
        pasta_app / "Eletrocentros.exe"
    ]
    for exe in exe_candidatos:
        if exe.exists():
            return [str(exe)]

    # 2. Script Python main.py
    main_py = pasta_app / "main.py"
    if main_py.exists():
        py_bin = obter_python_exe()
        return [py_bin, str(main_py)]

    return []


def executar_app(pasta_app: Path):
    """Inicia o processo do aplicativo em segundo plano e encerra o Launcher."""
    cmd = encontrar_executavel(pasta_app)
    if not cmd:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "Erro de Inicialização",
            f"Não foi possível localizar o executável ou main.py na pasta:\n{pasta_app}\n\nEntre em contato com o suporte de TI/Processos."
        )
        sys.exit(1)

    print(f"[Launcher] Iniciando aplicação: {' '.join(cmd)}")
    
    # Inicia desanexado para fechar o launcher imediatamente
    if sys.platform == "win32":
        creationflags = subprocess.DETACHED_PROCESS | subprocess.CREATE_NEW_PROCESS_GROUP
        subprocess.Popen(cmd, cwd=str(pasta_app), creationflags=creationflags, close_fds=True)
    else:
        subprocess.Popen(cmd, cwd=str(pasta_app))

    sys.exit(0)


def main():
    try:
        releases_dir = obter_pasta_releases_rede()
        manifest = ler_manifesto_rede(releases_dir)
        versao_local = ler_versao_local()

        print(f"[Launcher] Diretório de Releases: {releases_dir}")
        print(f"[Launcher] Versão Local em Cache: {versao_local or 'Nenhuma'}")

        if manifest is not None:
            versao_remota = manifest.get("version")
            pasta_remota = manifest.get("pasta", f"app-{versao_remota}")
            pasta_alvo = LOCAL_CACHE_DIR / pasta_remota

            # Atualiza se a versão for diferente ou se os arquivos locais ainda não existirem
            if versao_local != versao_remota or not pasta_alvo.exists():
                print(f"[Launcher] Nova versão detectada ({versao_remota}). Iniciando atualização...")
                pasta_alvo = atualizar_aplicativo(releases_dir, manifest)
                limpar_versoes_antigas(manter_ultimas=2)
            else:
                print("[Launcher] O aplicativo já está na versão mais recente.")

            executar_app(pasta_alvo)

        else:
            # Modo Offline / Falha de Rede: roda a última versão em cache
            print("[Launcher] Aviso: Rede indisponível. Tentando inicializar a partir do cache local...")
            if versao_local:
                pasta_alvo = LOCAL_CACHE_DIR / f"app-{versao_local}"
                if pasta_alvo.exists():
                    executar_app(pasta_alvo)

            # Se não há nada no cache local, tenta rodar da pasta base atual
            if (BASE_DIR / "main.py").exists():
                executar_app(BASE_DIR)

            # Erro fatal se não há nada para executar
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror(
                "Conexão Indisponível",
                "Não foi possível conectar à pasta de rede da WEG e não há nenhuma versão instalada no computador.\n\nPor favor, conecte-se à rede corporativa/VPN e tente novamente."
            )
            sys.exit(1)

    except Exception as e:
        import traceback
        err_msg = f"Ocorreu um erro inesperado ao iniciar o aplicativo:\n\n{str(e)}\n\nDetalhes:\n{traceback.format_exc()}"
        print(f"[Launcher Erro Fatal] {err_msg}")
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Erro no Launcher", err_msg)
        sys.exit(1)


if __name__ == "__main__":
    main()
