"""
Módulo de Extração e Parser Inteligente de PDF — Eletrocentros App
Processa Folhas de Dados de Eletrocentro WEG e especificações PCP
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
    Analisa a Folha de Dados do Eletrocentro e retorna um dicionário de campos reconhecidos
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
    texto_lower = texto.lower()

    # =========================================================================
    # 1. DADOS DE INSTALAÇÃO & CABEÇALHO (Cliente, PEP / Proposta)
    # =========================================================================
    cliente_match = re.search(r'CLIENTE(?:/PROJETO)?[:\s]+([^\n\r;]{3,70})', texto, re.IGNORECASE)
    if cliente_match:
        cli_val = cliente_match.group(1).strip()
        cli_val = re.sub(r'[\t\r\n]+', ' ', cli_val).strip(' -:')
        if cli_val and not cli_val.lower().startswith(('cidade', 'pais', 'país', 'ambiente')):
            campos_extraidos["cliente"] = cli_val
            logs_reconhecimento.append(f"Cliente/Projeto: {cli_val}")

    prop_match = re.search(r'N[°ºo]?\s*DA\s*PROPOSTA[:\s]+([^\n\r]{3,35})', texto, re.IGNORECASE)
    if prop_match:
        prop_val = prop_match.group(1).strip()
        campos_extraidos["proposta"] = prop_val
        logs_reconhecimento.append(f"Nº Proposta: {prop_val}")
        # Se parecer com código PEP, preenche também o campo PEP
        if re.search(r'^[A-Z0-9\-\.\/_]{6,25}$', prop_val, re.IGNORECASE):
            campos_extraidos["pep"] = prop_val

    # Busca PEP explícito caso exista
    pep_match = re.search(r'\b([A-Z0-9]{3}[-\.][A-Z0-9]{2}[-\.][A-Z0-9]{4,6}[-\.][A-Z0-9]{2,4})\b', texto, re.IGNORECASE)
    if not pep_match:
        pep_match = re.search(r'PEP[:\s]+([A-Z0-9\-\.\/_]{6,25})', texto, re.IGNORECASE)
    if pep_match:
        pep_val = pep_match.group(1).strip().upper()
        campos_extraidos["pep"] = pep_val
        logs_reconhecimento.append(f"PEP: {pep_val}")

    # =========================================================================
    # 2. TIPO DE ESTRUTURA (TIPO DE ELETROCENTRO)
    # Regra: Tipo de eletrocentro para definir o tipo de estrutura
    # =========================================================================
    tipo_eletro_match = re.search(r'TIPO\s*DE\s*ELETROCENTRO[:\s]+([^\n\r;]+?)(?:\s+MODELO|\s+FORMA|\s*$)', texto, re.IGNORECASE)
    if tipo_eletro_match:
        tipo_str = tipo_eletro_match.group(1).strip()
        tipo_lower = tipo_str.lower()
        if "embarcado" in tipo_lower:
            campos_extraidos["tipoestrutura"] = "Embarcado"
        elif "fixo" in tipo_lower:
            campos_extraidos["tipoestrutura"] = "Fixo"
        elif "semimóvel" in tipo_lower or "semimovel" in tipo_lower:
            campos_extraidos["tipoestrutura"] = "Semimóvel"
        elif "móvel" in tipo_lower or "movel" in tipo_lower:
            campos_extraidos["tipoestrutura"] = "Móvel"
        elif "modular" in tipo_lower:
            campos_extraidos["tipoestrutura"] = "Modular"
        elif "solar" in tipo_lower or "container" in tipo_lower:
            campos_extraidos["tipoestrutura"] = "Container Solar"
        elif "skid" in tipo_lower:
            campos_extraidos["tipoestrutura"] = "Skid (mecânica)"
        elif "essw" in tipo_lower:
            campos_extraidos["tipoestrutura"] = "ESSW (mecânica)"
        elif "pilotis" in tipo_lower:
            campos_extraidos["tipoestrutura"] = "Pilotis"
        else:
            campos_extraidos["tipoestrutura"] = tipo_str
        logs_reconhecimento.append(f"Tipo de Estrutura: {campos_extraidos['tipoestrutura']}")
    else:
        # Fallback de busca geral no texto
        if "embarcado" in texto_lower:
            campos_extraidos["tipoestrutura"] = "Embarcado"
        elif "container solar" in texto_lower:
            campos_extraidos["tipoestrutura"] = "Container Solar"
        elif "semimóvel" in texto_lower or "semimovel" in texto_lower:
            campos_extraidos["tipoestrutura"] = "Semimóvel"
        elif "móvel" in texto_lower or "movel" in texto_lower:
            campos_extraidos["tipoestrutura"] = "Móvel"
        elif "modular" in texto_lower:
            campos_extraidos["tipoestrutura"] = "Modular"
        elif "fixo" in texto_lower:
            campos_extraidos["tipoestrutura"] = "Fixo"

    # =========================================================================
    # 3. QUANTIDADE DE MÓDULOS NO COMPRIMENTO
    # Regra: pegar qtd modulos no comprimento para definir a quantidade de modulos
    # =========================================================================
    nmod = 1
    nmod_comp_match = re.search(r'QTDE\.?\s*M[ÓO]DULOS\s*NO\s*COMPRIMENTO[:\s]*(\d+)', texto, re.IGNORECASE)
    if not nmod_comp_match:
        nmod_comp_match = re.search(r'QUANTIDADE\s*DE\s*M[ÓO]DULOS\s*NO\s*COMPRIMENTO[:\s]*(\d+)', texto, re.IGNORECASE)
    if not nmod_comp_match:
        nmod_comp_match = re.search(r'(?:qtd\.?\s*módulos|n[ºo°]?\s*módulos?)[:\s]*(\d+)\b', texto, re.IGNORECASE)

    if nmod_comp_match:
        try:
            nmod = max(1, min(8, int(nmod_comp_match.group(1))))
        except Exception:
            nmod = 1

    campos_extraidos["nrmodulos"] = str(nmod)
    logs_reconhecimento.append(f"Qtd. Módulos no Comprimento: {nmod}")

    # =========================================================================
    # 4. DIMENSÕES DOS MÓDULOS (COMPRIMENTO mm, LARGURA mm)
    # Regra: as dimensoes colocar de acordo com a quantidade de modulos
    # =========================================================================
    comp_mm = 0.0
    larg_mm = 0.0

    comp_match = re.search(r'COMPRIMENTO\s*\(mm\)[:\s]*(\d+[\.,]?\d*)', texto, re.IGNORECASE)
    if comp_match:
        try:
            comp_mm = float(comp_match.group(1).replace(',', '.'))
        except Exception:
            comp_mm = 0.0

    larg_match = re.search(r'LARGURA\s*\(mm\)[:\s]*(\d+[\.,]?\d*)', texto, re.IGNORECASE)
    if larg_match:
        try:
            larg_mm = float(larg_match.group(1).replace(',', '.'))
        except Exception:
            larg_mm = 0.0

    if comp_mm > 0 or larg_mm > 0:
        # Se comprimento total informado em mm, divide pelos módulos no comprimento
        comp_por_modulo_m = round((comp_mm / nmod) / 1000.0, 2) if comp_mm > 0 else 10.0
        larg_por_modulo_m = round(larg_mm / 1000.0, 2) if larg_mm > 0 else 3.6

        modulos_list = []
        for i in range(nmod):
            modulos_list.append({
                "c": comp_por_modulo_m,
                "l": larg_por_modulo_m
            })
        campos_extraidos["modulos"] = modulos_list
        dim_str = f"{nmod}x ({comp_por_modulo_m}m x {larg_por_modulo_m}m) [Total: {round(comp_mm/1000.0, 2)}m x {round(larg_mm/1000.0, 2)}m]"
        logs_reconhecimento.append(f"Dimensões por Módulo: {dim_str}")

    # =========================================================================
    # 5. PÉ DIREITO (mm)
    # Regra: se o pé direito for maior que 3299mm então marca pe direito 3,3m
    # =========================================================================
    pe_match = re.search(r'P[ÉE]\s*DIREITO\s*\(mm\)[:\s]*(\d+[\.,]?\d*)', texto, re.IGNORECASE)
    if pe_match:
        try:
            pe_val_mm = float(pe_match.group(1).replace(',', '.'))
            if pe_val_mm > 3299:
                campos_extraidos["peDireito"] = True
                logs_reconhecimento.append(f"Pé Direito: {int(pe_val_mm)}mm (> 3299mm -> Pé Direito 3,3m: SIM)")
            else:
                campos_extraidos["peDireito"] = False
                logs_reconhecimento.append(f"Pé Direito: {int(pe_val_mm)}mm (<= 3299mm -> Pé Direito 3,3m: NÃO)")
        except Exception:
            pass

    # =========================================================================
    # 6. PISO / CHAPA REMOVÍVEL
    # =========================================================================
    piso_match = re.search(r'PISO.*?FORMA\s*CONSTRUTIVA[:\s]+([^\n\r;]{3,35})', texto, re.IGNORECASE | re.DOTALL)
    if piso_match:
        forma_piso = piso_match.group(1)
        if "removível" in forma_piso.lower() or "removivel" in forma_piso.lower():
            campos_extraidos["chapaRemovivel"] = True
            logs_reconhecimento.append("Chapa Removível: Sim")
        else:
            campos_extraidos["chapaRemovivel"] = False
            logs_reconhecimento.append("Chapa Removível: Não")
    elif "removível" in texto_lower or "removivel" in texto_lower:
        campos_extraidos["chapaRemovivel"] = True
        logs_reconhecimento.append("Chapa Removível: Sim")

    # =========================================================================
    # 7. PLANO DE PINTURA
    # =========================================================================
    pintura_match = re.search(r'BASE\s*E\s*ESTRUTURA\s*\(EM\s*A[ÇC]O\s*CARBONO\)[:\s]+([^\n\r;]{3,50})', texto, re.IGNORECASE)
    if pintura_match:
        pp_raw = pintura_match.group(1).strip()
        match_code = re.search(r'(WAU-[A-Z0-9\-]+|C[1-5]-?[MI]?|WEG\s*[0-9]+[A-Z]?)', pp_raw, re.IGNORECASE)
        if match_code:
            pp_clean = match_code.group(1).upper()
            campos_extraidos["planpin"] = pp_clean
            logs_reconhecimento.append(f"Plano de Pintura: {pp_clean}")
        else:
            campos_extraidos["planpin"] = pp_raw[:25]
            logs_reconhecimento.append(f"Plano de Pintura: {pp_raw[:25]}")
    else:
        for pp_term in ["WAU-ELETRO-08", "WAU-ELETRO-07", "WAU-ELETRO-06", "WAU-ELETRO-05", "C5-M", "C5-I", "C5", "C4", "C3"]:
            if pp_term.lower() in texto_lower:
                campos_extraidos["planpin"] = pp_term
                logs_reconhecimento.append(f"Plano de Pintura: {pp_term}")
                break

    # =========================================================================
    # 8. CLIMATIZAÇÃO / AR CONDICIONADO
    # =========================================================================
    hvac_tipo_match = re.search(r'CLIMATIZA[ÇC][ÃA]O.*?TIPO[:\s]+([^\n\r;]{3,35})', texto, re.IGNORECASE | re.DOTALL)
    tipo_hvac_raw = hvac_tipo_match.group(1).strip() if hvac_tipo_match else ""

    if "roof top" in tipo_hvac_raw.lower() or "rooftop" in tipo_hvac_raw.lower() or "roof top" in texto_lower or "rooftop" in texto_lower:
        campos_extraidos["tipomaq"] = "Roof Top"
        logs_reconhecimento.append("Ar Condicionado: Roof Top")
    elif "wall mounted" in tipo_hvac_raw.lower() or "wall-mounted" in tipo_hvac_raw.lower() or "wall mounted" in texto_lower:
        campos_extraidos["tipomaq"] = "Wall Mounted"
        logs_reconhecimento.append("Ar Condicionado: Wall Mounted")
    elif "self" in tipo_hvac_raw.lower() or "self" in texto_lower:
        campos_extraidos["tipomaq"] = "Self + Dutos"
        logs_reconhecimento.append("Ar Condicionado: Self + Dutos")
    elif "split" in tipo_hvac_raw.lower() or "splitão" in texto_lower or "split" in texto_lower:
        campos_extraidos["tipomaq"] = "Split Hi-Wall"
        logs_reconhecimento.append("Ar Condicionado: Split Hi-Wall")

    hvac_qtd_match = re.search(r'QTDE\.?\s*M[ÁA]QUINAS[:\s]*(\d+)', texto, re.IGNORECASE)
    if hvac_qtd_match:
        try:
            qtd_ac = int(hvac_qtd_match.group(1))
            campos_extraidos["qtdmaq"] = str(qtd_ac)
            logs_reconhecimento.append(f"Qtd. Ar Condicionado: {qtd_ac}")
        except Exception:
            pass

    # =========================================================================
    # 9. SISTEMAS (SEGURANÇA, CONTROLE DE ACESSO, CFTV, INCÊNDIO)
    # =========================================================================
    acesso_match = re.search(r'CONTROLE\s*DE\s*ACESSO.*?FORNECIMENTO[:\s]+([^\n\r;]{3,35})', texto, re.IGNORECASE | re.DOTALL)
    has_acesso = False
    if acesso_match:
        val_acesso = acesso_match.group(1).lower()
        if "sim" in val_acesso and not ("não" in val_acesso or "nao" in val_acesso or "sem" in val_acesso):
            has_acesso = True

    cftv_match = re.search(r'CFTV.*?FORNECIMENTO[:\s]+([^\n\r;]{3,35})', texto, re.IGNORECASE | re.DOTALL)
    has_cftv = False
    if cftv_match:
        val_cftv = cftv_match.group(1).lower()
        if "sim" in val_cftv and not ("não" in val_cftv or "nao" in val_cftv or "sem" in val_cftv):
            has_cftv = True

    if has_cftv and has_acesso:
        campos_extraidos["seguranca"] = "CFTV + Controle de Acesso"
        logs_reconhecimento.append("Sist. Segurança: CFTV + Controle de Acesso")
    elif has_cftv:
        campos_extraidos["seguranca"] = "CFTV"
        logs_reconhecimento.append("Sist. Segurança: CFTV")
    elif has_acesso:
        campos_extraidos["seguranca"] = "Controle de Acesso"
        logs_reconhecimento.append("Sist. Segurança: Controle de Acesso")
    else:
        campos_extraidos["seguranca"] = "Não possui"
        logs_reconhecimento.append("Sist. Segurança: Não possui")

    # Incêndio
    incendio_fornec = re.search(r'INC[ÊE]NDIO.*?FORNECIMENTO[:\s]+([^\n\r;]{3,40})', texto, re.IGNORECASE | re.DOTALL)
    if incendio_fornec and "não" in incendio_fornec.group(1).lower():
        campos_extraidos["incendio"] = "Não possui"
        logs_reconhecimento.append("Sist. Incêndio: Não possui")
    else:
        if any(k in texto_lower for k in ["fk-5-1-12", "novec 1230", "novec"]):
            campos_extraidos["incendio"] = "Novec 1230"
            logs_reconhecimento.append("Sist. Incêndio: Novec 1230 (FK-5-1-12)")
        elif any(k in texto_lower for k in ["fm-200", "fm200", "hfc-227"]):
            campos_extraidos["incendio"] = "FM-200"
            logs_reconhecimento.append("Sist. Incêndio: FM-200")
        elif "co2" in texto_lower and "extintor co2" not in texto_lower:
            campos_extraidos["incendio"] = "CO2"
            logs_reconhecimento.append("Sist. Incêndio: CO2")
        elif "aerosol" in texto_lower or "aerossol" in texto_lower:
            campos_extraidos["incendio"] = "Aerossol"
            logs_reconhecimento.append("Sist. Incêndio: Aerossol")
        elif "detecção" in texto_lower or "deteccao" in texto_lower or "central endereçada" in texto_lower:
            campos_extraidos["incendio"] = "Detecção e Alarme"
            logs_reconhecimento.append("Sist. Incêndio: Detecção e Alarme")

    # =========================================================================
    # 10. ACESSÓRIOS IDENTIFICADOS
    # =========================================================================
    acessorios_flags = []

    # Pilotis
    if re.search(r'PILOTIS.*?FORNECIMENTO[:\s]+Sim', texto, re.IGNORECASE | re.DOTALL):
        acessorios_flags.append("pilotis")
    # Escadas / Plataforma
    if re.search(r'ESCADA\s*DE\s*ACESSO[:\s]+Sim', texto, re.IGNORECASE) or re.search(r'PLATAFORMA\s*EXTERNA[:\s]+Sim', texto, re.IGNORECASE):
        acessorios_flags.append("esc_plat_padao_weg")
    # Porão de Cabos
    porao_match = re.search(r'POR[ÃA]O\s*DE\s*CABOS.*?TIPO[:\s]+([^\n\r;]{3,40})', texto, re.IGNORECASE | re.DOTALL)
    if porao_match and not ("sem porão" in porao_match.group(1).lower() or "sem porao" in porao_match.group(1).lower() or "não" in porao_match.group(1).lower()):
        acessorios_flags.append("porao_de_cabos")
    # Rede de Dutos
    if re.search(r'REDE\s*DE\s*DUTOS[:\s]+Sim', texto, re.IGNORECASE):
        acessorios_flags.append("rede_de_dutos")
    # Calhas Pluviais
    if re.search(r'SISTEMA\s*PLUVIAL.*?TIPO[:\s]+Com Pingadeiras', texto, re.IGNORECASE | re.DOTALL):
        acessorios_flags.append("calhas_pluviais")
    # Dutos de Gases
    if "dutos de gases" in texto_lower or "duto de gases" in texto_lower:
        acessorios_flags.append("duto_de_gases")

    if acessorios_flags:
        campos_extraidos["acessorios"] = acessorios_flags
        logs_reconhecimento.append(f"Acessórios ({len(acessorios_flags)}): {', '.join(acessorios_flags)}")

    # =========================================================================
    # 11. OUTROS PARÂMETROS TÉCNICOS
    # =========================================================================
    # Transformador a óleo c/ bacia
    if "baia de trafos" in texto_lower or "bacia de contenção" in texto_lower or "base skid para transformadores" in texto_lower:
        campos_extraidos["trafoOleo"] = True
        logs_reconhecimento.append("Transformador a Óleo / Bacia: Sim")

    # Teste de software / estanquiedade
    if "teste de estanquiedade" in texto_lower or "teste de software" in texto_lower or "teste de sw" in texto_lower:
        campos_extraidos["testesw"] = True
        logs_reconhecimento.append("Teste Especial / Software: Sim")

    return {
        "status": "success",
        "filename": filename,
        "total_campos": len(campos_extraidos),
        "fields": campos_extraidos,
        "logs": logs_reconhecimento,
        "resumo": f"{len(campos_extraidos)} parâmetro(s) extraído(s) com sucesso de '{filename}'."
    }
