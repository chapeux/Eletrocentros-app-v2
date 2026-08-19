"""
Módulo de Extração e Parser Inteligente de PDF — Eletrocentros App
Processa documentos técnicos, folhas de dados, propostas comerciais e especificações PCP
para preenchimento automático do formulário de planejamento.
"""

import re
import os
import io
import base64
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple


def extrair_texto_pdf(pdf_input: bytes | str | Path) -> str:
    """
    Extrai todo o conteúdo textual do arquivo PDF usando PyMuPDF (fitz), pdfplumber ou pypdfium2.
    """
    texto_completo = []

    # 1. Tentativa com PyMuPDF (fitz) - mais rápido e preciso
    try:
        import fitz
        if isinstance(pdf_input, (str, Path)) and os.path.exists(str(pdf_input)):
            doc = fitz.open(str(pdf_input))
        else:
            raw_bytes = pdf_input if isinstance(pdf_input, bytes) else pdf_input.encode('latin1')
            doc = fitz.open(stream=raw_bytes, filetype="pdf")
        
        for pagina in doc:
            texto_completo.append(pagina.get_text() or "")
        doc.close()
        
        resultado = "\n".join(texto_completo).strip()
        if resultado:
            return resultado
    except Exception as e:
        print(f"[PDF Parser] PyMuPDF não disponível ou erro ({e}), tentando pdfplumber...")

    # 2. Tentativa com pdfplumber
    try:
        import pdfplumber
        if isinstance(pdf_input, (str, Path)) and os.path.exists(str(pdf_input)):
            with pdfplumber.open(str(pdf_input)) as pdf:
                for p in pdf.pages:
                    texto_completo.append(p.extract_text() or "")
        else:
            raw_bytes = pdf_input if isinstance(pdf_input, bytes) else pdf_input.encode('latin1')
            with pdfplumber.open(io.BytesIO(raw_bytes)) as pdf:
                for p in pdf.pages:
                    texto_completo.append(p.extract_text() or "")

        resultado = "\n".join(texto_completo).strip()
        if resultado:
            return resultado
    except Exception as e:
        print(f"[PDF Parser] pdfplumber erro ({e}), tentando pypdfium2...")

    # 3. Tentativa com pypdfium2
    try:
        import pypdfium2 as pdfium
        if isinstance(pdf_input, (str, Path)) and os.path.exists(str(pdf_input)):
            pdf = pdfium.PdfDocument(str(pdf_input))
        else:
            raw_bytes = pdf_input if isinstance(pdf_input, bytes) else pdf_input.encode('latin1')
            pdf = pdfium.PdfDocument(raw_bytes)

        for page in pdf:
            textpage = page.get_textpage()
            texto_completo.append(textpage.get_text_range() or "")
            textpage.close()
            page.close()
        pdf.close()

        return "\n".join(texto_completo).strip()
    except Exception as e:
        print(f"[PDF Parser] Erro fatal na extração de texto PDF: {e}")

    return "\n".join(texto_completo).strip()


def parse_pdf_document(pdf_bytes_or_base64: bytes | str, filename: str = "documento.pdf") -> Dict[str, Any]:
    """
    Analisa o texto do PDF e retorna um dicionário de campos reconhecidos
    compatível com o formulário de Planejamento do Eletrocentros App.
    """
    if isinstance(pdf_bytes_or_base64, str):
        raw_str = pdf_bytes_or_base64
        if "," in raw_str:
            raw_str = raw_str.split(",", 1)[1]
        try:
            pdf_bytes = base64.b64decode(raw_str)
        except Exception:
            pdf_bytes = raw_str.encode('utf-8')
    else:
        pdf_bytes = pdf_bytes_or_base64

    texto = extrair_texto_pdf(pdf_bytes)
    if not texto:
        return {
            "status": "error",
            "message": "Não foi possível extrair texto do documento PDF. Verifique se o arquivo não é uma imagem digitalizada sem OCR."
        }

    campos_extraidos: Dict[str, Any] = {}
    logs_reconhecimento: List[str] = []

    # Normalização para matching (mantendo também texto original)
    texto_lower = texto.lower()

    # 1. PEP (Padrão WEG: XXX-XX-XXXXX-XXX ou variações alfanuméricas)
    pep_match = re.search(r'\b([A-Z0-9]{3}[-\.][A-Z0-9]{2}[-\.][A-Z0-9]{4,6}[-\.][A-Z0-9]{2,4})\b', texto, re.IGNORECASE)
    if not pep_match:
        pep_match = re.search(r'PEP[:\s]+([A-Z0-9\-\.\/_]{6,25})', texto, re.IGNORECASE)
    if pep_match:
        pep_val = pep_match.group(1).strip().upper()
        campos_extraidos["pep"] = pep_val
        logs_reconhecimento.append(f"PEP: {pep_val}")

    # 2. CLIENTE / RAZÃO SOCIAL
    cliente_match = re.search(r'(?:Cliente|Razão Social|Customer|Destinatário)[:\s]+([^\n\r;]{3,50})', texto, re.IGNORECASE)
    if cliente_match:
        cli_val = cliente_match.group(1).strip()
        # Remove caracteres indesejados
        cli_val = re.sub(r'[\t\r\n]+', ' ', cli_val).strip(' -:')
        if len(cli_val) > 2 and not cli_val.lower().startswith(('data', 'pep', 'cnpj', 'item')):
            campos_extraidos["cliente"] = cli_val
            logs_reconhecimento.append(f"Cliente: {cli_val}")

    # 3. TIPO DE ESTRUTURA
    # Lista de tipos conhecidos
    tipos_conhecidos = [
        ("Container Solar", ["container solar", "maritimo solar", "marítimo solar", "fotovoltaic", "solar container"]),
        ("Skid (mecânica)", ["skid mecânica", "skid mecanica", "skid mec"]),
        ("Skid (com elétrica)", ["skid com elétrica", "skid com eletrica", "skid eletrico", "skid elétrico"]),
        ("ESSW (mecânica)", ["essw mecânica", "essw mecanica", "bess mec"]),
        ("ESSW (elétrica)", ["essw elétrica", "essw eletrica", "bess ele"]),
        ("Pilotis", ["pilotis", "elevada sobre pilotis"]),
        ("Móvel", ["estrutura móvel", "estrutura movel", "eletrocentro móvel", "eletrocentro movel"]),
        ("Semimóvel", ["semimóvel", "semimovel", "semi-móvel", "semi-movel"]),
        ("Modular", ["modular", "multimodular", "multi-modular"]),
        ("Fixo", ["estrutura fixa", "eletrocentro fixo", "sala fixa"]),
        ("Embarcado", ["embarcado", "offshore", "naval"]),
        ("Eletrocentro", ["eletrocentro", "sala elétrica", "sala eletrica", "e-house", "shelter"])
    ]
    for tipo_nome, keywords in tipos_conhecidos:
        if any(k in texto_lower for k in keywords):
            campos_extraidos["tipoestrutura"] = tipo_nome
            logs_reconhecimento.append(f"Tipo de Estrutura: {tipo_nome}")
            break

    # 4. QUANTIDADE DE MÓDULOS
    nmod_match = re.search(r'(?:quantidade de módulos|qtd\.?\s*módulos|n[ºo°]?\s*módulos?|nmod|módulos?)[:\s]*(\d+)\b', texto, re.IGNORECASE)
    if not nmod_match:
        nmod_match = re.search(r'\b(\d+)\s*(?:módulos|modulos)\b', texto, re.IGNORECASE)
    
    nmod_val = 1
    if nmod_match:
        try:
            nmod_val = max(1, min(8, int(nmod_match.group(1))))
            campos_extraidos["nrmodulos"] = str(nmod_val)
            logs_reconhecimento.append(f"Qtd Módulos: {nmod_val}")
        except Exception:
            pass

    # 5. PLANO DE PINTURA
    planpin_match = re.search(r'(?:Plano de Pintura|Pintura|Plano Pintura)[:\s]*([A-Z0-9\-\. ]{2,20})', texto, re.IGNORECASE)
    if planpin_match:
        pp_val = planpin_match.group(1).strip()
        campos_extraidos["planpin"] = pp_val
        logs_reconhecimento.append(f"Plano de Pintura: {pp_val}")
    else:
        # Busca termos comuns de pintura (C3, C4, C5-M, C5-I, etc.)
        for pp_term in ["C5-M", "C5-I", "C5", "C4", "C3", "WEG 207A", "WEG 207", "WEG 203"]:
            if pp_term.lower() in texto_lower:
                campos_extraidos["planpin"] = pp_term
                logs_reconhecimento.append(f"Plano de Pintura: {pp_term}")
                break

    # 6. DIMENSÕES DOS MÓDULOS (Comprimento e Largura)
    # Padrões do tipo "12.0 x 3.6m", "10000 x 3200", "12m x 3,6m"
    dimensoes_encontradas: List[Tuple[float, float]] = []

    dim_matches = re.findall(r'(\d+[,\.]?\d*)\s*(?:m|mm|metros)?\s*[xX×]\s*(\d+[,\.]?\d*)\s*(?:m|mm|metros)?', texto)
    for c_raw, l_raw in dim_matches:
        try:
            c_num = float(c_raw.replace(',', '.'))
            l_num = float(l_raw.replace(',', '.'))
            # Converte mm para metros se necessário
            if c_num > 100: c_num /= 1000.0
            if l_num > 100: l_num /= 1000.0
            # Valida limites plausíveis de módulo de eletrocentro (ex: 2m a 30m comp, 1.5m a 6m larg)
            if 2.0 <= c_num <= 35.0 and 1.5 <= l_num <= 8.0:
                dimensoes_encontradas.append((round(c_num, 2), round(l_num, 2)))
        except Exception:
            continue

    if dimensoes_encontradas:
        modulos_list = []
        for i in range(nmod_val):
            dim_item = dimensoes_encontradas[i] if i < len(dimensoes_encontradas) else dimensoes_encontradas[0]
            modulos_list.append({"c": dim_item[0], "l": dim_item[1]})
        campos_extraidos["modulos"] = modulos_list
        dim_str = ", ".join([f"M{idx+1}: {m['c']}m x {m['l']}m" for idx, m in enumerate(modulos_list)])
        logs_reconhecimento.append(f"Dimensões: {dim_str}")

    # 7. AR CONDICIONADO / CLIMATIZAÇÃO
    if "roof top" in texto_lower or "rooftop" in texto_lower:
        campos_extraidos["tipomaq"] = "Roof Top"
        logs_reconhecimento.append("Ar Condicionado: Roof Top")
    elif "self" in texto_lower:
        campos_extraidos["tipomaq"] = "Self + Dutos"
        logs_reconhecimento.append("Ar Condicionado: Self + Dutos")
    elif "wall mounted" in texto_lower or "wall-mounted" in texto_lower or "wallmounted" in texto_lower:
        campos_extraidos["tipomaq"] = "Wall Mounted"
        logs_reconhecimento.append("Ar Condicionado: Wall Mounted")
    elif "split" in texto_lower:
        campos_extraidos["tipomaq"] = "Split Hi-Wall"
        logs_reconhecimento.append("Ar Condicionado: Split")

    # Quantidade de máquinas de AC
    qtd_ac_match = re.search(r'(\d+)\s*(?:máquinas|maquinas|unidades|unids|aparelhos|splits|rooftops)\s*(?:de\s*)?(?:ar|climatiz)', texto, re.IGNORECASE)
    if qtd_ac_match:
        try:
            qtd_ac = int(qtd_ac_match.group(1))
            campos_extraidos["qtdmaq"] = str(qtd_ac)
            logs_reconhecimento.append(f"Qtd Ar Condicionado: {qtd_ac}")
        except Exception:
            pass

    # 8. SISTEMA DE INCÊNDIO
    if any(k in texto_lower for k in ["novec", "novec 1230", "fk-5-1-12"]):
        campos_extraidos["incendio"] = "Novec 1230"
        logs_reconhecimento.append("Sist. Incêndio: Novec 1230")
    elif any(k in texto_lower for k in ["fm-200", "fm200", "hfc-227"]):
        campos_extraidos["incendio"] = "FM-200"
        logs_reconhecimento.append("Sist. Incêndio: FM-200")
    elif "aerosol" in texto_lower or "aerossol" in texto_lower:
        campos_extraidos["incendio"] = "Aerossol"
        logs_reconhecimento.append("Sist. Incêndio: Aerossol")
    elif "co2" in texto_lower:
        campos_extraidos["incendio"] = "CO2"
        logs_reconhecimento.append("Sist. Incêndio: CO2")
    elif "detecção e alarme" in texto_lower or "deteccao e alarme" in texto_lower or "sdae" in texto_lower:
        campos_extraidos["incendio"] = "Detecção e Alarme"
        logs_reconhecimento.append("Sist. Incêndio: Detecção e Alarme")

    # 9. SISTEMA DE SEGURANÇA
    if "cftv" in texto_lower and ("controle de acesso" in texto_lower or "acesso" in texto_lower):
        campos_extraidos["seguranca"] = "CFTV + Controle de Acesso"
        logs_reconhecimento.append("Sist. Segurança: CFTV + Controle de Acesso")
    elif "cftv" in texto_lower:
        campos_extraidos["seguranca"] = "CFTV"
        logs_reconhecimento.append("Sist. Segurança: CFTV")
    elif "controle de acesso" in texto_lower:
        campos_extraidos["seguranca"] = "Controle de Acesso"
        logs_reconhecimento.append("Sist. Segurança: Controle de Acesso")

    # 10. EQUIPAMENTOS / COLUNAS
    col_match = re.search(r'(\d+)\s*(?:colunas|cubículos|cubiculos|painéis|paineis|coluna\b)', texto, re.IGNORECASE)
    if col_match:
        try:
            qtd_col = int(col_match.group(1))
            if 1 <= qtd_col <= 100:
                campos_extraidos["nrcolunas"] = str(qtd_col)
                logs_reconhecimento.append(f"Nº Colunas Total: {qtd_col}")
        except Exception:
            pass

    # Complexidade
    if "complexidade alta" in texto_lower:
        campos_extraidos["complexidade"] = "Alta"
    elif "complexidade média" in texto_lower or "complexidade media" in texto_lower:
        campos_extraidos["complexidade"] = "Média"
    elif "complexidade baixa" in texto_lower:
        campos_extraidos["complexidade"] = "Baixa"

    # Transformador a óleo c/ bacia
    if any(k in texto_lower for k in ["bacia de contenção", "bacia de contencao", "transformador a óleo", "trafo a oleo", "trafo a óleo"]):
        campos_extraidos["trafoOleo"] = True
        logs_reconhecimento.append("Transformador a Óleo: Sim")

    # Teste de software
    if any(k in texto_lower for k in ["teste de software", "teste de sw", "teste integrado de sw"]):
        campos_extraidos["testesw"] = True
        logs_reconhecimento.append("Teste de Software: Sim")

    # Pé direito alto
    pe_match = re.search(r'pé direito[:\s]*(\d+[,\.]?\d*)', texto, re.IGNORECASE)
    if pe_match:
        try:
            alt_pe = float(pe_match.group(1).replace(',', '.'))
            if alt_pe >= 3.3:
                campos_extraidos["peDireito"] = True
                logs_reconhecimento.append("Pé Direito 3,3m: Sim")
        except Exception:
            pass

    # Chapa removível
    if "chapa removível" in texto_lower or "chapa removivel" in texto_lower:
        campos_extraidos["chapaRemovivel"] = True
        logs_reconhecimento.append("Chapa Removível: Sim")

    # 11. ACESSÓRIOS IDENTIFICADOS
    acessorios_flags = []
    if "escada" in texto_lower or "plataforma" in texto_lower:
        acessorios_flags.append("esc_plat_padao_weg")
    if "porão de cabos" in texto_lower or "porao de cabos" in texto_lower:
        acessorios_flags.append("porao_de_cabos")
    if "pilotis" in texto_lower:
        acessorios_flags.append("pilotis")
    if "rede de dutos" in texto_lower or "dutos de ar" in texto_lower:
        acessorios_flags.append("rede_de_dutos")
    if "fundo falso" in texto_lower:
        acessorios_flags.append("fundo_falso")
    if "calha pluvial" in texto_lower or "calhas pluviais" in texto_lower:
        acessorios_flags.append("calhas_pluviais")
    if "duto de gases" in texto_lower or "dutos de gases" in texto_lower:
        acessorios_flags.append("duto_de_gases")

    if acessorios_flags:
        campos_extraidos["acessorios"] = acessorios_flags
        logs_reconhecimento.append(f"Acessórios detectados: {len(acessorios_flags)}")

    # 12. CONTAINER SOLAR ESPECÍFICO
    if "programação de relés" in texto_lower or "reles de protecao" in texto_lower:
        campos_extraidos["progReles"] = True
    if "diagrama bti" in texto_lower:
        campos_extraidos["diagBTI"] = True
    if "diagrama agrupador" in texto_lower or "agrupamento" in texto_lower:
        campos_extraidos["diagAgrup"] = True

    return {
        "status": "success",
        "filename": filename,
        "total_campos": len(campos_extraidos),
        "fields": campos_extraidos,
        "logs": logs_reconhecimento,
        "resumo": f"{len(campos_extraidos)} parâmetro(s) extraído(s) com sucesso de '{filename}'."
    }
