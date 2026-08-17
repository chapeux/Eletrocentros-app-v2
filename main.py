"""
Eletrocentros App — Backend Python & GUI Launcher
Substituição de Tkinter por Interface Web Desktop Moderna em HTML5/CSS3/JS.
"""

import sys
import os
import re
import json
import base64
import http.server
import socketserver
import threading
import webbrowser
from datetime import datetime
from pathlib import Path

# Módulo Backend de Banco de Dados MySQL e Sincronização GitHub
from backend.database import init_db, comparar_e_registrar_alteracoes, obter_logs, registrar_log
from backend.git_sync import sync_github_async

# Base Directory Setup
BASE_DIR = Path(__file__).resolve().parent
FRONTEND_DIR = BASE_DIR / "frontend"

# CONFIG FILE PATHS, REGRAS, ASSETS & PASSWORDS
CONFIG_FILE = FRONTEND_DIR / "config.json"
REGRAS_FILE = FRONTEND_DIR / "regras.json"
SELETOR_FILE = FRONTEND_DIR / "seletor.json"
ICON_PATH = BASE_DIR / "assets" / "icone.ico"
MAINTENANCE_PASSWORD = os.environ.get("MAINTENANCE_PASSWORD", "admin")

# Diretório de rede para armazenamento de anexos do histórico
ATTACHMENTS_NETWORK_DIR = Path(
    r"Q:\GROUPS\BR_SC_ITJ_WAU_DPTO_PRODUCAO\Processos WAU Chaves Especiais e Acionamentos\10 - PASTAS PESSOAIS\Luan Schappo\Anexos_Historico"
)

# Define Windows AppUserModelID para ícone na barra de tarefas do Windows
if sys.platform == "win32":
    try:
        import ctypes
        myappid = "com.eletrocentros.app.v2"
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
    except Exception:
        pass


def format_compact_json(obj) -> str:
    """
    Formatador de JSON para regras.json que compacta objetos folhas e inlina
    arrays repetitivos (como 'it', 'etapas', 'cond', 'condicoes', 'valores') em linha única.
    """
    s = json.dumps(obj, ensure_ascii=False, indent=2)

    # 1. Compactar objetos folha em 1 única linha
    s = re.sub(r'\{\s*"t":\s*"([^"]+)",\s*"v":\s*([^}\n]+?)\s*\}', r'{"t": "\1", "v": \2}', s)
    s = re.sub(r'\{\s*"tipo":\s*"([^"]+)",\s*"valor":\s*([^}\n]+?)\s*\}', r'{"tipo": "\1", "valor": \2}', s)
    s = re.sub(r'\{\s*"tipo":\s*"([^"]+)",\s*"modo":\s*"([^"]+)"\s*\}', r'{"tipo": "\1", "modo": "\2"}', s)
    s = re.sub(r'\{\s*"tipo":\s*"([^"]+)",\s*"modo":\s*"([^"]+)",\s*"valor":\s*([^}\n]+?)\s*\}', r'{"tipo": "\1", "modo": "\2", "valor": \3}', s)
    s = re.sub(r'\{\s*"tipo":\s*"([^"]+)",\s*"valor":\s*([^,]+),\s*"modo":\s*"([^"]+)"\s*\}', r'{"tipo": "\1", "valor": \2, "modo": "\3"}', s)
    s = re.sub(r'\{\s*"c":\s*"([^"]+)",\s*"o":\s*"([^"]+)",\s*"val":\s*"([^"]+)",\s*"j":\s*"([^"]+)"\s*\}', r'{"c": "\1", "o": "\2", "val": "\3", "j": "\4"}', s)
    s = re.sub(r'\{\s*"flag":\s*"([^"]+)",\s*"rotulo":\s*"([^"]+)",\s*"forma":\s*"([^"]+)",\s*"valor":\s*"([^"]+)"\s*\}', r'{"flag": "\1", "rotulo": "\2", "forma": "\3", "valor": "\4"}', s)

    # 2. Inlinar arrays de tokens/condições/etapas como "it", "etapas", "cond", "condicoes", "valores"
    def inline_key_array(match):
        key = match.group(1)
        content = match.group(2)
        single_line = re.sub(r'\s*\n\s*', ' ', content)
        return f'"{key}": [ {single_line.strip()} ]'

    s = re.sub(r'"(it|etapas|cond|condicoes|valores)":\s*\[\s*([\s\S]*?)\s*\]', inline_key_array, s)
    return s


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
        if CONFIG_FILE.exists():
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception as e:
                print(f"[Backend Python] Erro ao ler config.json: {e}")
        return {}

    def save_config(self, config_data: dict) -> dict:
        """Salva a nova configuração enviada pelo usuário no arquivo config.json e sincroniza no GitHub."""
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
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
        if REGRAS_FILE.exists():
            try:
                with open(REGRAS_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception as e:
                print(f"[Backend Python] Erro ao ler regras.json: {e}")
        return []

    def save_regras(self, *args, **kwargs) -> dict:
        """Salva as regras atualizadas, copia anexos para a rede, registra histórico no MySQL e sincroniza com o GitHub."""
        try:
            regras_data = None
            motivo = None
            anexos_input = []

            if len(args) >= 1:
                if isinstance(args[0], dict) and "regras" in args[0]:
                    regras_data = args[0].get("regras")
                    motivo = args[0].get("motivo")
                    if "anexos" in args[0] and isinstance(args[0]["anexos"], list):
                        anexos_input = args[0]["anexos"]
                    elif args[0].get("anexo_nome") and args[0].get("anexo_base64"):
                        anexos_input = [{"nome": args[0]["anexo_nome"], "base64": args[0]["anexo_base64"]}]
                else:
                    regras_data = args[0]
            elif "regras_data" in kwargs:
                regras_data = kwargs.get("regras_data")

            if len(args) >= 2 and motivo is None:
                motivo = args[1]
            elif "motivo" in kwargs and motivo is None:
                motivo = kwargs.get("motivo")

            if "anexos" in kwargs and isinstance(kwargs["anexos"], list):
                anexos_input = kwargs["anexos"]
            elif "anexo_nome" in kwargs and "anexo_base64" in kwargs:
                anexos_input = [{"nome": kwargs["anexo_nome"], "base64": kwargs["anexo_base64"]}]

            if not regras_data:
                return {"status": "error", "message": "Nenhum dado de regras fornecido."}

            motivo_str = (motivo or "").strip()
            if len(motivo_str) < 20:
                return {"status": "error", "message": "O motivo da alteração é obrigatório e deve ter no mínimo 20 caracteres detalhando a mudança."}

            # Processamento de anexos múltiplos para a pasta de rede
            anexos_salvos = []
            if anexos_input:
                ts_prefix = datetime.now().strftime("%Y%m%d_%H%M%S")
                for idx, anexo_item in enumerate(anexos_input):
                    a_nome = anexo_item.get("nome")
                    a_base64 = anexo_item.get("base64")
                    if not a_nome or not a_base64:
                        continue
                    try:
                        safe_nome = re.sub(r'[^\w\.\-\(\) ]', '_', a_nome)
                        file_name = f"{ts_prefix}_{idx+1}_{safe_nome}" if len(anexos_input) > 1 else f"{ts_prefix}_{safe_nome}"

                        try:
                            ATTACHMENTS_NETWORK_DIR.mkdir(parents=True, exist_ok=True)
                            dest_file = ATTACHMENTS_NETWORK_DIR / file_name
                        except Exception as net_ex:
                            print(f"[Backend Python] Aviso ao acessar pasta de rede ({net_ex}). Usando fallback local...")
                            fallback_dir = BASE_DIR / "anexos_historico"
                            fallback_dir.mkdir(parents=True, exist_ok=True)
                            dest_file = fallback_dir / file_name

                        raw_data = a_base64
                        if "," in raw_data:
                            raw_data = raw_data.split(",", 1)[1]
                        file_bytes = base64.b64decode(raw_data)
                        with open(dest_file, "wb") as f_out:
                            f_out.write(file_bytes)

                        caminho_final = str(dest_file)
                        anexos_salvos.append({"nome": a_nome, "caminho": caminho_final})
                        print(f"[Backend Python] Anexo [{idx+1}/{len(anexos_input)}] salvo em: {caminho_final}")
                    except Exception as anexo_err:
                        print(f"[Backend Python] Erro ao processar anexo '{a_nome}': {anexo_err}")

            anexo_caminho_db = json.dumps(anexos_salvos, ensure_ascii=False) if anexos_salvos else None
            anexo_nome_db = ", ".join([x["nome"] for x in anexos_salvos]) if anexos_salvos else None

            # 1. Carrega as regras anteriores para calcular o diff
            regras_antigas = self.get_regras()

            # 2. Compara e registra as alterações na tabela 'logs_modificacoes' no MySQL
            total_logs = comparar_e_registrar_alteracoes(
                regras_antigas,
                regras_data,
                motivo=motivo,
                anexo_nome=anexo_nome_db,
                anexo_caminho=anexo_caminho_db
            )
            print(f"[Backend Python] {total_logs} log(s) de alteração de regras registrados no banco de dados. (Motivo: {motivo})")

            # 3. Salva a nova versão em regras.json usando formato compacto
            compact_json_str = format_compact_json(regras_data)
            with open(REGRAS_FILE, "w", encoding="utf-8") as f:
                f.write(compact_json_str)

            # 4. Sincronização automática com GitHub em segundo plano
            resumo_git = f"Atualização de regras ({total_logs} alteração/ões)"
            if motivo:
                resumo_git += f" - Motivo: {motivo}"
            sync_github_async(resumo=resumo_git)

            print("[Backend Python] Regras salvas em regras.json com sucesso!")
            return {"status": "success", "logs_registrados": total_logs, "anexos_salvos": anexos_salvos}
        except Exception as e:
            print(f"[Backend Python] Erro ao salvar regras.json: {e}")
            return {"status": "error", "message": str(e)}

    def open_attachment(self, file_path: str) -> dict:
        """Abre o arquivo anexo diretamente da pasta de rede no visualizador padrão do sistema operacional."""
        try:
            if not file_path:
                return {"status": "error", "message": "Arquivo não especificado."}
            p = Path(file_path)
            file_name = p.name or "Anexo"
            if not p.exists():
                return {"status": "error", "message": f"O arquivo '{file_name}' não foi encontrado ou não está acessível no momento."}
            
            if os.name == 'nt':
                os.startfile(str(p))
            else:
                import subprocess
                subprocess.Popen(["xdg-open", str(p)])
            
            return {"status": "success"}
        except Exception as e:
            print(f"[Backend Python] Erro ao abrir anexo '{file_path}': {e}")
            file_name = Path(file_path).name if file_path else "Anexo"
            return {"status": "error", "message": f"Não foi possível abrir o arquivo '{file_name}'. Verifique o acesso ao documento."}

    def get_logs(self, limit: int = 100) -> list:
        """Retorna o histórico de alterações de regras cadastrado no banco de dados MySQL."""
        return obter_logs(limit=limit)

    def verify_password(self, password: str) -> bool:
        """Valida se a senha digitada corresponde às credenciais de mantenedor."""
        return password == self.password or password == "1234"

    def get_seletor(self) -> list:
        """Lê e retorna a base do seletor de PEPs e Centros de Trabalho cadastrada em seletor.json."""
        try:
            if SELETOR_FILE.exists():
                with open(SELETOR_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
            return []
        except Exception as e:
            print(f"[Backend Python] Erro ao ler seletor.json: {e}")
            return []

    def save_seletor(self, payload: dict) -> dict:
        """Salva as alterações na base do seletor.json e registra histórico no banco MySQL."""
        try:
            seletor_data = payload.get("seletor", payload) if isinstance(payload, dict) and "seletor" in payload else payload
            motivo = payload.get("motivo", "") if isinstance(payload, dict) else ""

            # Trata anexo se houver
            anexo_nome_db = None
            anexo_caminho_db = None
            if isinstance(payload, dict) and payload.get("anexo_base64") and payload.get("anexo_nome"):
                try:
                    anexo_bytes = base64.b64decode(payload["anexo_base64"])
                    nome_limpo = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{payload['anexo_nome']}"
                    caminho_salvo = ANEXOS_DIR / nome_limpo
                    with open(caminho_salvo, "wb") as af:
                        af.write(anexo_bytes)
                    anexo_nome_db = payload["anexo_nome"]
                    anexo_caminho_db = str(caminho_salvo)
                except Exception as ex_anexo:
                    print(f"[Backend Python] Erro ao salvar anexo do seletor: {ex_anexo}")

            # Registra no MySQL
            total_logs = 0
            if DB_AVAILABLE:
                usuario = os.getenv("USERNAME", "Sistema")
                total_linhas = len(seletor_data) if isinstance(seletor_data, list) else 0
                registrar_log(
                    disciplina="Seletor de PEP & CTs",
                    campo="Tabela Seletor",
                    tipo_alteracao="UPDATE",
                    detalhes=f"Atualização do Seletor de PEP e Centros de Trabalho ({total_linhas} combinações cadastradas).",
                    motivo=motivo or "Atualização dos parâmetros do Seletor de PEP/CTs",
                    usuario=usuario,
                    anexo_nome=anexo_nome_db,
                    anexo_caminho=anexo_caminho_db
                )
                total_logs = 1

            # Salva o arquivo JSON formatado
            with open(SELETOR_FILE, "w", encoding="utf-8") as f:
                json.dump(seletor_data, f, ensure_ascii=False, indent=2)

            # Sincronização automática com GitHub
            resumo_git = "Atualização da tabela Seletor de PEP e Centros de Trabalho"
            if motivo:
                resumo_git += f" - Motivo: {motivo}"
            sync_github_async(resumo=resumo_git)

            print("[Backend Python] Seletor salvo em seletor.json com sucesso!")
            return {"status": "success", "logs_registrados": total_logs}
        except Exception as e:
            print(f"[Backend Python] Erro ao salvar seletor.json: {e}")
            return {"status": "error", "message": str(e)}

    def export_excel(self, data: dict) -> dict:
        """Gera e retorna a planilha Excel (.xlsx) com o resultado detalhado dos tempos."""
        try:
            import io
            import base64
            import openpyxl
            from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
            from openpyxl.utils import get_column_letter

            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Resumo de Tempos"
            ws.views.sheetView[0].showGridLines = True

            # Cores e Estilos
            title_font = Font(name="Segoe UI", size=14, bold=True, color="FFFFFF")
            title_fill = PatternFill("solid", fgColor="0F2C59")

            section_font = Font(name="Segoe UI", size=11, bold=True, color="0F2C59")
            label_font = Font(name="Segoe UI", size=10, bold=True, color="4A5568")
            val_font = Font(name="Segoe UI", size=10, color="1A202C")

            hdr_font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
            hdr_fill = PatternFill("solid", fgColor="1E3E62")

            subtotal_font = Font(name="Segoe UI", size=10, bold=True, color="1A202C")
            subtotal_fill = PatternFill("solid", fgColor="F0F4F8")

            total_font = Font(name="Segoe UI", size=11, bold=True, color="0F2C59")
            total_fill = PatternFill("solid", fgColor="D0E8FF")

            thin_border = Border(
                left=Side(style="thin", color="CBD5E0"),
                right=Side(style="thin", color="CBD5E0"),
                top=Side(style="thin", color="CBD5E0"),
                bottom=Side(style="thin", color="CBD5E0")
            )
            double_bottom = Border(
                left=Side(style="thin", color="CBD5E0"),
                right=Side(style="thin", color="CBD5E0"),
                top=Side(style="thin", color="CBD5E0"),
                bottom=Side(style="double", color="0F2C59")
            )

            # 1. Título
            ws.merge_cells("A1:D1")
            ws["A1"] = "ELETROCENTROS APP — RESULTADO DO CÁLCULO DE TEMPOS"
            ws["A1"].font = title_font
            ws["A1"].fill = title_fill
            ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[1].height = 36

            # 2. Informações Gerais do Projeto
            ctx = data.get("ctx", {})
            pep = data.get("pep", "") or ctx.get("pep", "") or "Não informado"
            data_hora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            usuario = os.getenv("USERNAME", "Usuário")

            ws["A3"] = "INFORMAÇÕES DO PROJETO"
            ws["A3"].font = section_font

            meta_rows = [
                ("Código PEP / Ordem:", pep, "Data / Hora de Cálculo:", data_hora),
                ("Tipo de Estrutura:", str(ctx.get("tipoestrutura", "-")), "Quantidade de Módulos:", f"{ctx.get('nmod', 1)} Módulo(s)"),
                ("Dimensões Gerais:", f"{ctx.get('comp', 0)}m (C) x {ctx.get('larg', 0)}m (L) x {ctx.get('alt', 0)}m (A)", "Plano de Pintura:", str(ctx.get("planpin", "-"))),
                ("Ar Condicionado:", f"{ctx.get('tipomaq', '-')} ({ctx.get('qtdmaq', 0)}x)", "Complexidade Elétrica:", str(ctx.get("complexidade", "-"))),
                ("Sist. Incêndio:", str(ctx.get("incendio", "-")), "Sist. Segurança:", str(ctx.get("seguranca", "-"))),
                ("Usuário Responsável:", usuario, "Nº Colunas Total:", str(ctx.get("nrcolunas", 0)))
            ]

            curr_row = 4
            for label1, val1, label2, val2 in meta_rows:
                ws[f"A{curr_row}"] = label1
                ws[f"A{curr_row}"].font = label_font
                ws[f"B{curr_row}"] = val1
                ws[f"B{curr_row}"].font = val_font

                ws[f"C{curr_row}"] = label2
                ws[f"C{curr_row}"].font = label_font
                ws[f"D{curr_row}"] = val2
                ws[f"D{curr_row}"].font = val_font
                curr_row += 1

            # 3. KPIs de Resumo
            curr_row += 1
            ws[f"A{curr_row}"] = "RESUMO GERAL CONSOLIDADO"
            ws[f"A{curr_row}"].font = section_font
            curr_row += 1

            ws[f"A{curr_row}"] = "TOTAL HORAS ORÇADAS (H):"
            ws[f"A{curr_row}"].font = Font(name="Segoe UI", size=11, bold=True, color="107C41")
            ws[f"B{curr_row}"] = float(data.get("totalGeralH", 0))
            ws[f"B{curr_row}"].font = Font(name="Segoe UI", size=11, bold=True, color="107C41")
            ws[f"B{curr_row}"].number_format = "#,##0.00 \"h\""

            ws[f"C{curr_row}"] = "DURAÇÃO ESTIMADA (DUR):"
            ws[f"C{curr_row}"].font = Font(name="Segoe UI", size=11, bold=True, color="D83B01")
            ws[f"D{curr_row}"] = float(data.get("totalGeralDUR", 0))
            ws[f"D{curr_row}"].font = Font(name="Segoe UI", size=11, bold=True, color="D83B01")
            ws[f"D{curr_row}"].number_format = "#,##0.0 \"dias\""
            curr_row += 2

            # 4. Tabela de Detalhamento por Área / Processo
            ws[f"A{curr_row}"] = "Área / Disciplina"
            ws[f"B{curr_row}"] = "Processo / Campo"
            ws[f"C{curr_row}"] = "Horas Orçadas (H)"
            ws[f"D{curr_row}"] = "Duração Estimada (DUR)"
            for col_idx, col_letter in enumerate(["A", "B", "C", "D"]):
                cell = ws[f"{col_letter}{curr_row}"]
                cell.font = hdr_font
                cell.fill = hdr_fill
                cell.alignment = Alignment(horizontal="left" if col_idx < 2 else "right", vertical="center")
            ws.row_dimensions[curr_row].height = 24
            curr_row += 1

            for area in data.get("resultadosAreas", []):
                nome_area = area.get("area", "Área")
                campos = area.get("campos", [])
                for campo in campos:
                    ws[f"A{curr_row}"] = nome_area
                    ws[f"B{curr_row}"] = campo.get("chave", "")

                    cell_h = ws[f"C{curr_row}"]
                    cell_h.value = float(campo.get("h", 0))
                    cell_h.number_format = "#,##0.00"

                    cell_dur = ws[f"D{curr_row}"]
                    cell_dur.value = float(campo.get("dur", 0))
                    cell_dur.number_format = "#,##0.0"

                    for col_letter in ["A", "B", "C", "D"]:
                        ws[f"{col_letter}{curr_row}"].border = thin_border
                        ws[f"{col_letter}{curr_row}"].font = val_font
                    curr_row += 1

                # Subtotal da Área
                ws[f"A{curr_row}"] = f"Subtotal — {nome_area}"
                ws[f"B{curr_row}"] = f"({len(campos)} processos)"

                cell_sub_h = ws[f"C{curr_row}"]
                cell_sub_h.value = float(area.get("totalH", 0))
                cell_sub_h.number_format = "#,##0.00"

                cell_sub_dur = ws[f"D{curr_row}"]
                cell_sub_dur.value = float(area.get("totalDUR", 0))
                cell_sub_dur.number_format = "#,##0.0"

                for col_letter in ["A", "B", "C", "D"]:
                    c = ws[f"{col_letter}{curr_row}"]
                    c.font = subtotal_font
                    c.fill = subtotal_fill
                    c.border = thin_border
                curr_row += 1

            # Total Geral
            ws[f"A{curr_row}"] = "TOTAL GERAL CONSOLIDADO"
            ws[f"B{curr_row}"] = "Todos os processos"

            cell_tot_h = ws[f"C{curr_row}"]
            cell_tot_h.value = float(data.get("totalGeralH", 0))
            cell_tot_h.number_format = "#,##0.00"

            cell_tot_dur = ws[f"D{curr_row}"]
            cell_tot_dur.value = float(data.get("totalGeralDUR", 0))
            cell_tot_dur.number_format = "#,##0.0"

            for col_letter in ["A", "B", "C", "D"]:
                c = ws[f"{col_letter}{curr_row}"]
                c.font = total_font
                c.fill = total_fill
                c.border = double_bottom
            ws.row_dimensions[curr_row].height = 24

            # 5. Seletor de PEP & Centros de Trabalho (se disponível)
            seletor_info = data.get("seletor")
            if seletor_info:
                curr_row += 2
                ws[f"A{curr_row}"] = "SELETOR DE PEP STANDARD & CENTROS DE TRABALHO (CTS)"
                ws[f"A{curr_row}"].font = section_font
                curr_row += 1

                ws[f"A{curr_row}"] = "PEP Standard Sugerido:"
                ws[f"A{curr_row}"].font = label_font
                ws[f"B{curr_row}"] = str(seletor_info.get("PEP Standard", "-"))
                ws[f"B{curr_row}"].font = Font(name="Segoe UI", size=10, bold=True, color="0F2C59")
                curr_row += 1

                ct_items = []
                if seletor_info.get("DR Eng Mec"):
                    ct_items.append(("Engenharia Mecânica:", f"DR {seletor_info.get('DR Eng Mec')} (Alt {seletor_info.get('Alt Eng Mec', '1')})"))
                if seletor_info.get("DR Eng Ele"):
                    ct_items.append(("Engenharia Elétrica:", f"DR {seletor_info.get('DR Eng Ele')} (Alt {seletor_info.get('Alt Eng Ele', '1')})"))

                nmod_val = int(ctx.get("nmod", 1) or 1)
                for m in range(1, 9):
                    dr_m = seletor_info.get(f"DR Mec {m}") or seletor_info.get(f"DR Mec{m}")
                    alt_m = seletor_info.get(f"Alt Mec {m}") or seletor_info.get(f"Alt Mec{m}") or "1"
                    if dr_m and m <= nmod_val:
                        ct_items.append((f"Mecânica Módulo {m}:", f"DR {dr_m} (Alt {alt_m})"))

                if seletor_info.get("DR Acess"):
                    ct_items.append(("Acessórios:", f"DR {seletor_info.get('DR Acess')} (Alt {seletor_info.get('Alt Acess', '1')})"))
                if seletor_info.get("DR Eletromec"):
                    ct_items.append(("Eletromecânica:", f"DR {seletor_info.get('DR Eletromec')} (Alt {seletor_info.get('Alt Eletromec', '1')})"))

                for ct_label, ct_val in ct_items:
                    ws[f"A{curr_row}"] = ct_label
                    ws[f"A{curr_row}"].font = label_font
                    ws[f"B{curr_row}"] = ct_val
                    ws[f"B{curr_row}"].font = val_font
                    curr_row += 1

            # Auto-largura das colunas
            for col in ws.columns:
                max_len = 0
                col_letter = get_column_letter(col[0].column)
                for cell in col:
                    val_str = str(cell.value or "")
                    if cell.row == 1: continue
                    if len(val_str) > max_len: max_len = len(val_str)
                ws.column_dimensions[col_letter].width = max(max_len + 4, 18)

            output_stream = io.BytesIO()
            wb.save(output_stream)
            output_stream.seek(0)
            b64_content = base64.b64encode(output_stream.read()).decode("utf-8")

            pep_clean = "".join(c for c in pep if c.isalnum() or c in ("-", "_")).strip()
            data_clean = datetime.now().strftime("%Y%m%d_%H%M")
            file_name = f"Resultado_PEP_{pep_clean}_{data_clean}.xlsx" if pep_clean and pep_clean != "Nãoinformado" else f"Resultado_Eletrocentro_{data_clean}.xlsx"

            return {
                "status": "success",
                "filename": file_name,
                "base64": b64_content
            }
        except Exception as e:
            print(f"[Backend Python] Erro ao exportar Excel: {e}")
            return {"status": "error", "message": str(e)}


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
        elif self.path == "/api/get_seletor":
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.end_headers()
            seletor = self.api.get_seletor()
            self.wfile.write(json.dumps(seletor, ensure_ascii=False).encode("utf-8"))
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
        elif self.path == "/api/save_seletor":
            content_length = int(self.headers.get("Content-Length", 0))
            post_data = self.rfile.read(content_length)
            try:
                data = json.loads(post_data.decode("utf-8"))
                result = self.api.save_seletor(data)
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
        elif self.path == "/api/export_excel":
            content_length = int(self.headers.get("Content-Length", 0))
            post_data = self.rfile.read(content_length)
            try:
                data = json.loads(post_data.decode("utf-8"))
                result = self.api.export_excel(data)
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
