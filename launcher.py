"""
Eletrocentros App — Launcher & Auto-Updater Inteligente (MySQL & Zero-Network-Share)
Verifica novas versões publicadas no Banco de Dados MySQL da WEG, sincroniza
com o cache local (%LOCALAPPDATA%\\EletrocentrosApp) e executa o aplicativo.
"""

import os
import sys
import json
import shutil
import hashlib
import zipfile
import subprocess
import time
from io import BytesIO
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox

# 1. Definição do diretório base e paths
if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).resolve().parent
    INTERNAL_DIR = Path(getattr(sys, "_MEIPASS", BASE_DIR))
else:
    BASE_DIR = Path(__file__).resolve().parent
    INTERNAL_DIR = BASE_DIR

LOCAL_APPDATA = Path(os.environ.get("LOCALAPPDATA", Path.home() / "AppData" / "Local"))
LOCAL_CACHE_DIR = LOCAL_APPDATA / "EletrocentrosApp"
CURRENT_VERSION_FILE = LOCAL_CACHE_DIR / "current_version.txt"

ICON_PATH = INTERNAL_DIR / "assets" / "icone.ico"
if not ICON_PATH.exists():
    ICON_PATH = BASE_DIR / "assets" / "icone.ico"

# Configuração do Banco de Dados MySQL
DB_CONFIG = {
    "host": os.environ.get("DB_HOST", "dcprd036187.weg.net"),
    "port": int(os.environ.get("DB_PORT", 8502)),
    "user": os.environ.get("DB_USER", "root"),
    "password": os.environ.get("DB_PASS", "neoempresarial"),
    "database": os.environ.get("DB_NAME", "bd_eletrocentros_app"),
    "connect_timeout": 4
}


def get_mysql_connection():
    """Tenta estabelecer conexão com o banco de dados MySQL com timeout seguro."""
    try:
        import mysql.connector
        return mysql.connector.connect(**DB_CONFIG)
    except Exception as e:
        print(f"[Launcher MySQL] Falha de conexão: {e}")
        return None


def buscar_ponteiro_release_mysql() -> dict | None:
    """Busca o ponteiro da release atual no MySQL (app_settings.release_atual)."""
    conn = get_mysql_connection()
    if conn is None:
        return None

    cursor = conn.cursor()
    try:
        cursor.execute("SELECT valor FROM app_settings WHERE chave = 'release_atual'")
        row = cursor.fetchone()
        if row is None:
            return None
        raw_val = row[0]
        return json.loads(raw_val) if isinstance(raw_val, str) else raw_val
    except Exception as e:
        print(f"[Launcher] Erro ao consultar app_settings.release_atual: {e}")
        return None
    finally:
        cursor.close()
        conn.close()


def baixar_pacote_mysql(version: str, splash: UpdateSplash | None = None) -> tuple[bytes, str]:
    """Baixa o pacote zip e o hash SHA-256 da tabela app_releases em chunks seguros."""
    conn = get_mysql_connection()
    if conn is None:
        raise RuntimeError("Não foi possível conectar ao banco de dados MySQL para baixar a release.")

    cursor = conn.cursor()
    try:
        cursor.execute(
            "SELECT sha256, tamanho_bytes FROM app_releases WHERE version = %s AND status = 'publicada'",
            (version,)
        )
        row = cursor.fetchone()
        if row is None:
            raise RuntimeError(f"Versão {version} não encontrada ou não está com status 'publicada'.")

        sha256_esperado, total_bytes = row[0], row[1]

        # Download em chunks de 8 MB (MySQL SUBSTRING é 1-indexed)
        chunk_size = 8 * 1024 * 1024
        chunks = []
        pos = 1

        while pos <= total_bytes:
            if splash:
                perc = min(100, int((pos / total_bytes) * 100))
                splash.set_status(f"Baixando pacote do servidor ({perc}%)...")
            cursor.execute(
                "SELECT SUBSTRING(pacote_zip, %s, %s) FROM app_releases WHERE version = %s",
                (pos, chunk_size, version)
            )
            part_row = cursor.fetchone()
            if not part_row or not part_row[0]:
                break
            part_data = part_row[0]
            chunks.append(part_data)
            pos += len(part_data)

        pacote_completo = b"".join(chunks)
        return pacote_completo, sha256_esperado
    finally:
        cursor.close()
        conn.close()


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
            text=f"Instalando a versão {versao_nova} via Banco de Dados WEG...",
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
            text="Conectando e baixando pacote do servidor...",
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


def aplicar_pacote_zip(pacote_bytes: bytes, version: str, splash: UpdateSplash | None = None) -> Path:
    """Extrai atomicamente o pacote ZIP baixado no cache local."""
    nome_pasta = f"app-{version}"
    pasta_alvo = LOCAL_CACHE_DIR / nome_pasta
    pasta_tmp = LOCAL_CACHE_DIR / f"{nome_pasta}.tmp"

    if pasta_tmp.exists():
        shutil.rmtree(pasta_tmp, ignore_errors=True)
    pasta_tmp.mkdir(parents=True, exist_ok=True)

    if splash:
        splash.set_status("Extraindo componentes do aplicativo...")

    with zipfile.ZipFile(BytesIO(pacote_bytes)) as zf:
        zf.extractall(pasta_tmp)

    if pasta_alvo.exists():
        shutil.rmtree(pasta_alvo, ignore_errors=True)

    pasta_tmp.rename(pasta_alvo)

    # Grava versão atual
    LOCAL_CACHE_DIR.mkdir(parents=True, exist_ok=True)
    CURRENT_VERSION_FILE.write_text(version, encoding="utf-8")

    return pasta_alvo


def obter_python_exe() -> str:
    """Retorna o caminho do interpretador Python para executar main.py."""
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
    
    if sys.platform == "win32":
        creationflags = subprocess.DETACHED_PROCESS | subprocess.CREATE_NEW_PROCESS_GROUP
        subprocess.Popen(cmd, cwd=str(pasta_app), creationflags=creationflags, close_fds=True)
    else:
        subprocess.Popen(cmd, cwd=str(pasta_app))

    sys.exit(0)


def main():
    try:
        versao_local = ler_versao_local()
        print(f"[Launcher] Versão Local em Cache: {versao_local or 'Nenhuma'}")

        # 1. Consulta o MySQL para obter a versão atual
        ponteiro = buscar_ponteiro_release_mysql()

        if ponteiro is not None:
            versao_remota = ponteiro.get("version")
            nome_pasta = f"app-{versao_remota}"
            pasta_alvo = LOCAL_CACHE_DIR / nome_pasta

            print(f"[Launcher] Versão no Banco MySQL: {versao_remota}")

            # Se a versão for diferente ou os arquivos locais não existirem, baixa do MySQL
            if versao_local != versao_remota or not pasta_alvo.exists():
                print(f"[Launcher] Nova versão detectada ({versao_remota}). Iniciando download via MySQL...")
                splash = UpdateSplash(versao_remota)

                try:
                    splash.set_status("Baixando pacote do servidor MySQL...")
                    pacote_bytes, sha_esperado = baixar_pacote_mysql(versao_remota, splash)

                    # Validação de integridade SHA-256
                    sha_calculado = hashlib.sha256(pacote_bytes).hexdigest()
                    if sha_calculado != sha_esperado:
                        print("[Launcher] Aviso: Hash inconsistente no primeiro download. Tentando novamente...")
                        splash.set_status("Reconectando para novo download...")
                        pacote_bytes, sha_esperado = baixar_pacote_mysql(versao_remota, splash)
                        sha_calculado = hashlib.sha256(pacote_bytes).hexdigest()
                        if sha_calculado != sha_esperado:
                            raise RuntimeError(f"Integridade inválida do pacote (Hash esperado: {sha_esperado}, calculado: {sha_calculado}).")

                    splash.set_status("Instalando e validando arquivos...")
                    pasta_alvo = aplicar_pacote_zip(pacote_bytes, versao_remota, splash)
                    limpar_versoes_antigas(manter_ultimas=2)
                    time.sleep(0.3)
                finally:
                    splash.close()
            else:
                print("[Launcher] O aplicativo já está na versão mais recente.")

            executar_app(pasta_alvo)

        else:
            # 2. Modo Offline / Falha de Conexão: executa a última versão em cache
            print("[Launcher] Aviso: Banco MySQL indisponível. Tentando inicializar do cache local...")
            if versao_local:
                pasta_alvo = LOCAL_CACHE_DIR / f"app-{versao_local}"
                if pasta_alvo.exists():
                    executar_app(pasta_alvo)

            # Fallback para execução direta da pasta base
            if (BASE_DIR / "main.py").exists():
                executar_app(BASE_DIR)

            root = tk.Tk()
            root.withdraw()
            messagebox.showerror(
                "Conexão Indisponível",
                "Não foi possível conectar ao servidor MySQL da WEG e não há nenhuma versão instalada no computador.\n\nPor favor, verifique a conexão de rede corporativa/VPN e tente novamente."
            )
            sys.exit(1)

    except Exception as e:
        import traceback
        err_msg = f"Ocorreu um erro ao verificar ou atualizar o aplicativo:\n\n{str(e)}\n\nDetalhes:\n{traceback.format_exc()}"
        print(f"[Launcher Erro] {err_msg}")
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Erro no Launcher", err_msg)
        sys.exit(1)


if __name__ == "__main__":
    main()
