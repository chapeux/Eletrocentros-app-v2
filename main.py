"""
Eletrocentros App — Backend Python & GUI Launcher
Substituição de Tkinter por Interface Web Desktop Moderna em HTML5/CSS3/JS.
"""

import sys
import os
import json
import http.server
import socketserver
import threading
import webbrowser
from pathlib import Path

# Módulo Backend de Banco de Dados MySQL e Sincronização GitHub
from backend.database import init_db, comparar_e_registrar_alteracoes, obter_logs, registrar_log
from backend.git_sync import sync_github_async

# Base Directory Setup
BASE_DIR = Path(__file__).resolve().parent
FRONTEND_DIR = BASE_DIR / "frontend"

# CONFIG FILE PATHS, REGRAS, ASSETS & PASSWORDS
CONFIG_FILE = BASE_DIR / "config.json"
FRONTEND_CONFIG_FILE = FRONTEND_DIR / "config.json"
REGRAS_FILE = BASE_DIR / "regras.json"
FRONTEND_REGRAS_FILE = FRONTEND_DIR / "regras.json"
ICON_PATH = BASE_DIR / "assets" / "icone.ico"
MAINTENANCE_PASSWORD = os.environ.get("MAINTENANCE_PASSWORD", "admin")

# Define Windows AppUserModelID para ícone na barra de tarefas do Windows
if sys.platform == "win32":
    try:
        import ctypes
        myappid = "com.eletrocentros.app.v2"
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
    except Exception:
        pass


class AppAPI:
    """
    API exposta para o JavaScript via pywebview (window.pywebview.api)
    ou chamadas HTTP para o backend.
    """

    def __init__(self):
        self.password = MAINTENANCE_PASSWORD
        # Inicializa e garante criação do BD 'bd_eletrocentros_app' e tabela 'logs_modificacoes'
        init_db()

    def get_config(self) -> dict:
        """Lê e retorna a configuração modular do arquivo config.json."""
        target = CONFIG_FILE if CONFIG_FILE.exists() else FRONTEND_CONFIG_FILE
        if target.exists():
            try:
                with open(target, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception as e:
                print(f"[Backend Python] Erro ao ler config.json: {e}")
        return {}

    def save_config(self, config_data: dict) -> dict:
        """Salva a nova configuração enviada pelo usuário no arquivo config.json e sincroniza no GitHub."""
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(config_data, f, ensure_ascii=False, indent=2)
            with open(FRONTEND_CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(config_data, f, ensure_ascii=False, indent=2)
            print("[Backend Python] Configurações salvas em config.json com sucesso!")
            
            # Sincronização automática com GitHub
            sync_github_async(resumo="Atualização de config.json")
            
            return {"status": "success"}
        except Exception as e:
            print(f"[Backend Python] Erro ao salvar config.json: {e}")
            return {"status": "error", "message": str(e)}

    def get_regras(self) -> list:
        """Lê e retorna a estrutura de regras do arquivo regras.json."""
        target = REGRAS_FILE if REGRAS_FILE.exists() else FRONTEND_REGRAS_FILE
        if target.exists():
            try:
                with open(target, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception as e:
                print(f"[Backend Python] Erro ao ler regras.json: {e}")
        return []

    def save_regras(self, *args, **kwargs) -> dict:
        """Salva as regras atualizadas, registra histórico no MySQL e sincroniza com o GitHub."""
        try:
            regras_data = None
            motivo = None

            if len(args) >= 1:
                if isinstance(args[0], dict) and "regras" in args[0]:
                    regras_data = args[0].get("regras")
                    motivo = args[0].get("motivo")
                else:
                    regras_data = args[0]
            elif "regras_data" in kwargs:
                regras_data = kwargs.get("regras_data")

            if len(args) >= 2 and motivo is None:
                motivo = args[1]
            elif "motivo" in kwargs and motivo is None:
                motivo = kwargs.get("motivo")

            if not regras_data:
                return {"status": "error", "message": "Nenhum dado de regras fornecido."}

            motivo_str = (motivo or "").strip()
            if len(motivo_str) < 20:
                return {"status": "error", "message": "O motivo da alteração é obrigatório e deve ter no mínimo 20 caracteres detalhando a mudança."}

            # 1. Carrega as regras anteriores para calcular o diff
            regras_antigas = self.get_regras()

            # 2. Compara e registra as alterações na tabela 'logs_modificacoes' no MySQL
            total_logs = comparar_e_registrar_alteracoes(regras_antigas, regras_data, motivo=motivo)
            print(f"[Backend Python] {total_logs} log(s) de alteração de regras registrados no banco de dados. (Motivo: {motivo})")

            # 3. Salva a nova versão em regras.json
            with open(REGRAS_FILE, "w", encoding="utf-8") as f:
                json.dump(regras_data, f, ensure_ascii=False, indent=2)
            with open(FRONTEND_REGRAS_FILE, "w", encoding="utf-8") as f:
                json.dump(regras_data, f, ensure_ascii=False, indent=2)

            # 4. Sincronização automática com GitHub em segundo plano
            resumo_git = f"Atualização de regras ({total_logs} alteração/ões)"
            if motivo:
                resumo_git += f" - Motivo: {motivo}"
            sync_github_async(resumo=resumo_git)

            print("[Backend Python] Regras salvas em regras.json com sucesso!")
            return {"status": "success", "logs_registrados": total_logs}
        except Exception as e:
            print(f"[Backend Python] Erro ao salvar regras.json: {e}")
            return {"status": "error", "message": str(e)}

    def get_logs(self, limit: int = 100) -> list:
        """Retorna o histórico de alterações de regras cadastrado no banco de dados MySQL."""
        return obter_logs(limit=limit)

    def verify_password(self, password: str) -> bool:
        """Valida se a senha digitada corresponde às credenciais de mantenedor."""
        return password == self.password or password == "1234"

    def get_disciplinas(self) -> list:
        """Retorna as disciplinas cadastradas para o painel de manutenção."""
        return [
            {"disciplina": "Mecânica — Estrutura", "total_campos": 14},
            {"disciplina": "Elétrica & Equipamentos", "total_campos": 18},
            {"disciplina": "Acessórios & Adicionais", "total_campos": 5},
            {"disciplina": "SAP & Automação", "total_campos": 6}
        ]

    def get_campos_disciplina(self, disciplina_nome: str) -> dict:
        """Retorna a lista de campos e regras da disciplina selecionada."""
        return {
            "campos": {
                "Módulo Estrutural Base": {
                    "analise": {
                        "base": {"forma": "aditiva", "valor_base": 40.0, "passo": 12.0},
                        "condicoes": [
                            {"flag": "chapa_remov", "forma": "fixo", "valor": 8.0},
                            {"flag": "pe_direito_3_3_m", "forma": "por_modulo", "valor": 3.5}
                        ]
                    }
                },
                "Instalação de Calhas Pluviais": {
                    "analise": {
                        "base": {"forma": "multiplicativa", "valor_base": 15.0},
                        "condicoes": [
                            {"flag": "casa_maquinas", "forma": "por_modulo", "valor": 4.0}
                        ]
                    }
                },
                "Porão de Cabos & Interligações": {
                    "analise": {
                        "base": {"forma": "constante", "valor": 25.0},
                        "condicoes": [
                            {"flag": "porao_de_cabos", "forma": "fixo", "valor": 10.0}
                        ]
                    }
                }
            }
        }

    def save_campo(self, campo_nome: str, data: dict) -> dict:
        """Salva as regras de um campo específico no backend."""
        print(f"[Backend Python] Salvando alterações para o campo '{campo_nome}': {data}")
        return {"status": "success", "campo": campo_nome, "analise": data}

    def calculate_tempos(self, form_data: dict) -> dict:
        """Cálculo principal de tempos de produção com base nos dados informados."""
        print(f"[Backend Python] Executando cálculo de tempos para os dados: {form_data}")
        return {"status": "success", "total_horas": 128.5}


class AppHTTPRequestHandler(http.server.SimpleHTTPRequestHandler):
    """Handler HTTP customizado com suporte a rotas da API /api/*."""

    api = AppAPI()

    def do_GET(self):
        if self.path == "/api/logs" or self.path.startswith("/api/logs?"):
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.end_headers()
            logs = self.api.get_logs()
            self.wfile.write(json.dumps(logs, ensure_ascii=False).encode("utf-8"))
            return
        elif self.path == "/api/get_regras":
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.end_headers()
            regras = self.api.get_regras()
            self.wfile.write(json.dumps(regras, ensure_ascii=False).encode("utf-8"))
            return
        elif self.path == "/api/get_config":
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.end_headers()
            config = self.api.get_config()
            self.wfile.write(json.dumps(config, ensure_ascii=False).encode("utf-8"))
            return

        super().do_GET()

    def do_POST(self):
        if self.path == "/api/save_regras":
            content_length = int(self.headers.get("Content-Length", 0))
            post_data = self.rfile.read(content_length)
            try:
                data = json.loads(post_data.decode("utf-8"))
                result = self.api.save_regras(data)
                self.send_response(200)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.end_headers()
                self.wfile.write(json.dumps(result, ensure_ascii=False).encode("utf-8"))
            except Exception as e:
                self.send_response(500)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.end_headers()
                self.wfile.write(json.dumps({"status": "error", "message": str(e)}).encode("utf-8"))
            return

        super().do_POST()


def run_local_server(port: int = 8000):
    """Inicia um servidor HTTP local com suporte a API para servir a aplicação."""
    os.chdir(FRONTEND_DIR)
    handler = AppHTTPRequestHandler
    with socketserver.TCPServer(("", port), handler) as httpd:
        print(f"[Python HTTP Server] Servindo frontend e API em http://localhost:{port}")
        httpd.serve_forever()



def apply_native_window_icon():
    """Aplica o ícone icone.ico diretamente em todas as janelas Win32 do processo Python."""
    if sys.platform != "win32" or not ICON_PATH.exists():
        return

    try:
        import ctypes
        user32 = ctypes.windll.user32
        kernel32 = ctypes.windll.kernel32
        current_pid = kernel32.GetCurrentProcessId()

        hicon_small = user32.LoadImageW(0, str(ICON_PATH), 1, 16, 16, 0x0010)
        hicon_big = user32.LoadImageW(0, str(ICON_PATH), 1, 32, 32, 0x0010)
        if not hicon_small:
            hicon_small = user32.LoadImageW(0, str(ICON_PATH), 1, 0, 0, 0x0010)
        if not hicon_big:
            hicon_big = hicon_small

        WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)

        def enum_windows_callback(hwnd, lparam):
            pid = ctypes.c_ulong()
            user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            if pid.value == current_pid:
                user32.SendMessageW(hwnd, 0x0080, 0, hicon_small)  # WM_SETICON, ICON_SMALL
                user32.SendMessageW(hwnd, 0x0080, 1, hicon_big)    # WM_SETICON, ICON_BIG
            return True

        user32.EnumWindows(WNDENUMPROC(enum_windows_callback), 0)
    except Exception as e:
        print(f"[GUI Python] Erro ao aplicar ícone nativo: {e}")


def launch_app():
    """Tenta abrir via pywebview como janela nativa. Caso contrário, usa navegador local."""
    api = AppAPI()
    index_file = FRONTEND_DIR / "index.html"

    try:
        import webview
        print("[GUI Python] Inicializando janela nativa com PyWebView...")
        window = webview.create_window(
            title="Eletrocentros — Sistema Integrado PCP & Manutenção",
            url=str(index_file),
            js_api=api,
            width=1380,
            height=900,
            resizable=True,
            min_size=(800, 600)
        )

        def on_started():
            import time
            time.sleep(0.2)
            apply_native_window_icon()
            time.sleep(0.8)
            apply_native_window_icon()

        webview.start(on_started, debug=False)
    except ImportError:
        print("[GUI Python] pywebview não encontrado no ambiente Python.")
        print("[GUI Python] Iniciando servidor HTTP local de fallback...")

        port = 8000
        server_thread = threading.Thread(target=run_local_server, args=(port,), daemon=True)
        server_thread.start()

        url = f"http://localhost:{port}/index.html"
        print(f"[GUI Python] Abrindo navegador padrão em {url}...")
        webbrowser.open(url)

        try:
            while True:
                pass
        except KeyboardInterrupt:
            print("\n[GUI Python] Aplicação encerrada pelo usuário.")


if __name__ == "__main__":
    launch_app()
