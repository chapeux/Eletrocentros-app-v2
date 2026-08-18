"""
Script de Publicação e Empacotamento de Releases — Eletrocentros App
Monta a pasta de distribuição `releases/app-<versao>` e atualiza o manifesto `releases/version.json`.
Pode ser executado com ou sem PyInstaller.
"""

import os
import sys
import json
import shutil
import subprocess
from pathlib import Path
from datetime import datetime

BASE_DIR = Path(__file__).resolve().parent.parent
RELEASES_DIR = BASE_DIR / "releases"
VERSION_JSON = RELEASES_DIR / "version.json"


def obter_versao_atual() -> tuple[str, str]:
    if VERSION_JSON.exists():
        try:
            data = json.loads(VERSION_JSON.read_text(encoding="utf-8"))
            return data.get("version", "2.0.0"), data.get("notas", "Atualização do sistema")
        except Exception:
            pass
    return "2.0.0", "Atualização do sistema"


def publicar_release(versao: str, notas: str, compilar_exe: bool = False):
    print("=" * 70)
    print(f" [BUILD RELEASE] Gerando Release {versao}")
    print("=" * 70)

    pasta_nome = f"app-{versao}"
    pasta_destino = RELEASES_DIR / pasta_nome
    pasta_tmp = RELEASES_DIR / f"{pasta_nome}.tmp"

    if pasta_tmp.exists():
        shutil.rmtree(pasta_tmp, ignore_errors=True)
    pasta_tmp.mkdir(parents=True, exist_ok=True)

    # 1. Copiar pastas e arquivos essenciais
    itens_para_copiar = ["backend", "frontend", "assets", "main.py", "requirements.txt"]
    for item_nome in itens_para_copiar:
        src = BASE_DIR / item_nome
        dst = pasta_tmp / item_nome
        if src.is_dir():
            print(f"  -> Copiando diretório: {item_nome}")
            shutil.copytree(src, dst, dirs_exist_ok=True)
        elif src.is_file():
            print(f"  -> Copiando arquivo: {item_nome}")
            shutil.copy2(src, dst)

    # 2. Compilar com PyInstaller se solicitado
    if compilar_exe:
        print("  -> Compilando executável com PyInstaller...")
        icone = BASE_DIR / "assets" / "icone.ico"
        cmd = [
            sys.executable, "-m", "PyInstaller",
            "--noconfirm",
            "--onedir",
            "--windowed",
            "--name", "EletrocentrosApp",
            f"--add-data", f"{BASE_DIR / 'frontend'};frontend",
            f"--add-data", f"{BASE_DIR / 'assets'};assets",
            f"--add-data", f"{BASE_DIR / 'backend'};backend",
            str(BASE_DIR / "main.py")
        ]
        if icone.exists():
            cmd.extend(["--icon", str(icone)])

        res = subprocess.run(cmd, cwd=str(BASE_DIR))
        if res.returncode == 0:
            dist_folder = BASE_DIR / "dist" / "EletrocentrosApp"
            if dist_folder.exists():
                print("  -> Copiando artefatos compilados do PyInstaller...")
                shutil.copytree(dist_folder, pasta_tmp, dirs_exist_ok=True)
        else:
            print("  [Aviso] Falha na compilação PyInstaller. Mantendo modo interpretado Python.")

    # 3. Mover atomicamente para a pasta final de release
    if pasta_destino.exists():
        shutil.rmtree(pasta_destino, ignore_errors=True)
    pasta_tmp.rename(pasta_destino)
    print(f"  [OK] Pasta de release pronta em: {pasta_destino}")

    # 4. Atualizar manifesto version.json
    manifest = {
        "version": versao,
        "pasta": pasta_nome,
        "notas": notas,
        "obrigatoria": True,
        "data_publicacao": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }
    with open(VERSION_JSON, "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)

    print(f"  [OK] Manifesto {VERSION_JSON.name} atualizado com sucesso!")
    print("=" * 70)


if __name__ == "__main__":
    v_atual, n_atual = obter_versao_atual()
    versao_target = sys.argv[1] if len(sys.argv) > 1 else v_atual
    notas_target = sys.argv[2] if len(sys.argv) > 2 else n_atual
    compilar = "--exe" in sys.argv

    publicar_release(versao_target, notas_target, compilar_exe=compilar)
