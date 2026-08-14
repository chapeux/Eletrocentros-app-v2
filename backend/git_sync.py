"""
Módulo de Sincronização Automática com o GitHub — Eletrocentros App
Realiza o commit e push automático de regras.json e config.json para o GitHub.
"""

import subprocess
import threading
import getpass
import os
from pathlib import Path
from datetime import datetime

BASE_DIR = Path(__file__).resolve().parent.parent


def get_current_user() -> str:
    try:
        user = getpass.getuser()
        if user:
            return user
    except Exception:
        pass
    return os.environ.get("USERNAME", os.environ.get("USER", "usuario_sistema"))


def _auto_git_push_worker(commit_msg: str):
    try:
        # Adiciona arquivos modificados
        subprocess.run(["git", "add", "frontend/regras.json", "frontend/config.json"], cwd=BASE_DIR, check=False)
        
        # Realiza o commit
        commit_res = subprocess.run(["git", "commit", "-m", commit_msg], cwd=BASE_DIR, capture_output=True, text=True, check=False)
        
        # Faz o push para o GitHub (branch main)
        push_res = subprocess.run(["git", "push", "origin", "main"], cwd=BASE_DIR, capture_output=True, text=True, check=False)
        
        if push_res.returncode == 0:
            print(f"[Git Sync Python] Push automático no GitHub concluído: '{commit_msg}'")
        else:
            print(f"[Git Sync Python] Status do push GitHub: {push_res.stdout or push_res.stderr}")
    except Exception as e:
        print(f"[Git Sync Python] Erro na sincronização automática com GitHub: {e}")


def sync_github_async(resumo: str = "Atualização de regras/configuração", usuario: str = None):
    """
    Dispara o commit e push automático para o repositório GitHub de forma assíncrona,
    sem travar a resposta da API ou a interface do usuário.
    """
    if not usuario:
        usuario = get_current_user()
    
    data_str = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    commit_msg = f"{resumo} por '{usuario}' em {data_str}"
    
    thread = threading.Thread(target=_auto_git_push_worker, args=(commit_msg,), daemon=True)
    thread.start()
