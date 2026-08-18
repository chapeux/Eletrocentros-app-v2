import json
import copy

def load_json(path):
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)

def classificar_disciplina_tarefa(tarefa_code, desc):
    """Classifica a fase/disciplina da tarefa pelo código e descrição."""
    try:
        t_num = int(str(tarefa_code).strip())
    except:
        t_num = 0

    desc_upper = desc.upper()
    if 500 <= t_num < 600:
        if any(w in desc_upper for w in ["LOM", "PBS", "PPA", "PCI", "PCE", "PAC", "LCA", "LAA", "LAM", "LMA", "PTR", "LMT"]):
            return "Engenharia Mecânica"
        if any(w in desc_upper for w in ["PIL", "PCL", "LMC", "PIN", "LMI", "PSS", "LMS", "DIN", "LMD", "PBA", "LMB", "PRF"]):
            return "Engenharia Elétrica"
        if any(w in desc_upper for w in ["ROM", "531", "551", "561", "581", "585", "589", "ROTEIRO"]):
            return "Processos"
        return "Engenharia"
    elif 600 <= t_num < 700:
        if any(w in desc_upper for w in ["ROM", "ROTEIRO"]):
            return "Processos"
        return "Engenharia Elétrica"
    elif 700 <= t_num < 800:
        if any(w in desc_upper for w in ["PIN", "PINTURA"]):
            return "Pintura"
        if any(w in desc_upper for w in ["COR", "FCH", "PRB", "SBA", "PRE", "SES", "EDF", "CHI", "CHE", "ESTRUTURA", "CALDEIRARIA"]):
            return "Mecânica"
        return "Montagem Mecânica"
    elif 800 <= t_num < 900:
        if any(w in desc_upper for w in ["LBA", "INT", "TES", "INS", "PEE", "PEM", "FEC", "FEA", "FEQ", "ELETR"]):
            return "Elétrica / Testes"
        return "Acessórios & Elétrica"
    elif 900 <= t_num:
        return "Faturamento & Encerramento"
    return "Geral"

def run_schedule_engine(form_data, calc_times, template_blocks, seletor_rows=None):
    """
    Gera o cronograma operacional completo fiel ao VBA original.
    """
    tipoestrutura = str(form_data.get('tipoestrutura', ''))
    betim = bool(form_data.get('Betim1310', False))
    sem_eng = bool(form_data.get('SemEng', False))
    nrmodulos_str = str(form_data.get('nrmodulos', '1 Módulo'))
    tipomaq = str(form_data.get('tipomaq', 'Não possui'))
    seguranca = str(form_data.get('seguranca', 'Não possui'))
    testesw = bool(form_data.get('testesw', False))
    chk_filho = bool(form_data.get('chkFilho', False))
    chaparemov = bool(form_data.get('chaparemov', False))
    trafooleo = bool(form_data.get('trafooleo', False))
    programacaoreles = bool(form_data.get('programacaoreles', False))
    pro_bti = bool(form_data.get('proBTI', False))
    
    # 1. Identificar cenário
    cenarios = template_blocks.get('cenarios', {})
    if tipoestrutura in ["Container Solar", "ESSW (mecânica)"]:
        cenario_id = "container_solar_essw_mecanica"
    elif tipoestrutura == "Skid (mecânica)":
        cenario_id = "skid_mecanica_com_betim" if betim else "skid_mecanica_sem_betim"
    elif tipoestrutura == "ESSW (elétrica)":
        cenario_id = "essw_eletrica"
    elif tipoestrutura == "Pilotis":
        cenario_id = "pilotis"
    elif tipoestrutura == "Skid (com elétrica)":
        cenario_id = "skid_com_eletrica_betim_true"
    elif tipoestrutura == "Serviço Engenharia":
        cenario_id = "servico_engenharia"
    else:
        cenario_id = "eletrocentro_padrao"

    cenario = cenarios.get(cenario_id)
    if not cenario:
        raise ValueError(f"Cenário desconhecido: {cenario_id}")

    tarefas_base = copy.deepcopy(cenario.get('tarefas', []))
    
    def get_val(k, default=0.0):
        v = calc_times.get(k, default)
        try:
            return float(v)
        except:
            return 0.0

    grid = {}
    for t in tarefas_base:
        r = t['linha_template'] + 2
        dur_val = t['duracao']
        try:
            dur_num = float(str(dur_val).strip().replace(',', '.'))
        except:
            dur_num = 0.0
        trab_val = t['trabalho']
        try:
            trab_num = float(str(trab_val).strip().replace(',', '.'))
        except:
            trab_num = 0.0
            
        grid[r] = {
            'row': r,
            'tarefa': t['tarefa'],
            'descricao_tarefa': t['descricao_tarefa'],
            'duracao': dur_num,
            'unidade': t['unidade'],
            'trabalho': trab_num,
            'calculado': False
        }

    def set_cell(r, h_key, d_key):
        if r in grid:
            if h_key in calc_times:
                grid[r]['trabalho'] = round(get_val(h_key), 1)
                grid[r]['calculado'] = True
            if d_key in calc_times:
                grid[r]['duracao'] = round(get_val(d_key), 1)
                grid[r]['calculado'] = True

    def delete_excel_rows(min_r, max_r):
        nonlocal grid
        to_del = [r for r in grid.keys() if min_r <= r <= max_r]
        for r in to_del:
            del grid[r]
        num_deleted = max_r - min_r + 1
        new_grid = {}
        for r in sorted(grid.keys()):
            if r < min_r:
                new_grid[r] = grid[r]
            else:
                new_r = r - num_deleted
                grid[r]['row'] = new_r
                new_grid[new_r] = grid[r]
        grid = new_grid

    # =========================================================================
    # CENÁRIO 1: CONTAINER SOLAR / ESSW MECÂNICA
    # =========================================================================
    if cenario_id == "container_solar_essw_mecanica":
        if tipoestrutura in ["Container Solar", "ESSW (mecânica)"]:
            if 10 in grid: grid[10]['descricao_tarefa'] = "PEC - Projeto Estrutura Container"
            if 12 in grid: grid[12]['descricao_tarefa'] = "EMC - Estagiamento Mat. Estrut. Cont."
            if 16 in grid: grid[16]['descricao_tarefa'] = "PEI - Projeto Estrutura Interna"
            if 18 in grid: grid[18]['descricao_tarefa'] = "EMI - Estagiamento Mat. Estr. Interna"
            if 69 in grid: grid[69]['descricao_tarefa'] = "FPC - Fabricação Peças Caldeiraria"
            if 74 in grid: grid[74]['descricao_tarefa'] = "OEE - Ordens Estrutura Interna"
            if 77 in grid: grid[77]['descricao_tarefa'] = "SEI - Separação Estrutura Interna"
            if 78 in grid: grid[78]['descricao_tarefa'] = "MEI - Montagem Estrutura Interna"
            if 121 in grid: grid[121]['descricao_tarefa'] = "OFE - Ordens Fechamento Externo"
            if 124 in grid: grid[124]['descricao_tarefa'] = "SAF - Separação Almox. Fech. Externo"
            if 126 in grid: grid[126]['descricao_tarefa'] = "SFE - Separação Fechamento Externo"
            if 127 in grid: grid[127]['descricao_tarefa'] = "MFE - Montagem Fechamento Externo"
            
        if chk_filho and 153 in grid:
            grid[153]['tarefa'] = 899
            
        if 131 in grid: grid[131]['descricao_tarefa'] = "SII - Separação Almox. Instal./ Inc."
        if 133 in grid: grid[133]['descricao_tarefa'] = "MII - Montagem Instalações / Incêndio"
        if seguranca != "Não possui" and 142 in grid:
            grid[142]['descricao_tarefa'] = "LBS - Leito e Bandejamento / Sist. Seg."

        set_cell(4, 'HorLOM', 'DurLOM')
        set_cell(8, 'HorLMM', 'DurLMM')
        set_cell(10, 'HorPBS', 'DurPBS')
        set_cell(13, 'HorPPA', 'DurPPA')
        set_cell(16, 'HorPCI', 'DurPCI')
        set_cell(19, 'HorPCE', 'DurPCE')
        set_cell(22, 'HorPAC', 'DurPAC')
        set_cell(24, 'HorLCA', 'DurLCA')
        set_cell(27, 'HorLAA', 'DurLAA')
        set_cell(30, 'HorLAM', 'DurLAM')
        set_cell(33, 'HorLMA', 'DurLMA')
        set_cell(37, 'HorPTR', 'DurPTR')
        set_cell(41, 'HorPIL', 'DurPIL')
        set_cell(45, 'HorPCL', 'DurPCL')
        set_cell(47, 'HorLMC', 'DurLMC')
        set_cell(49, 'HorPIN', 'DurPIN')
        set_cell(51, 'HorLMI', 'DurLMI')
        
        if seguranca != "Não possui":
            set_cell(53, 'HorPSS', 'DurPSS')
            set_cell(55, 'HorLMS', 'DurLMS')
            
        set_cell(43, 'HorLMT', 'DurLMT')
        set_cell(57, 'HorDIN', 'DurDIN')
        set_cell(59, 'HorLMD', 'DurLMD')
        set_cell(61, 'HorPBA', 'DurPBA')
        set_cell(63, 'HorLMB', 'DurLMB')
        set_cell(65, 'HorPRF', 'DurPRF')
        
        # Processos
        set_cell(11, 'Hor531', 'Dur531')
        set_cell(17, 'Hor551', 'Dur551')
        set_cell(20, 'Hor561', 'Dur561')
        set_cell(28, 'Hor581', 'Dur581')
        set_cell(31, 'Hor585', 'Dur585')
        set_cell(34, 'Hor589', 'Dur589')
        
        # Módulo 1
        set_cell(69, 'HorCOR1', 'DurCOR1')
        set_cell(70, 'HorFCH1', 'DurFCH1')
        set_cell(72, 'HorEDF1', 'DurEDF1')
        set_cell(75, 'HorPIN1', 'DurPIN1')
        set_cell(78, 'HorCHI1', 'DurCHI1')
        set_cell(71, 'HorCHE1', 'DurCHE1')
        
        # Acessórios
        set_cell(120, 'HorFAC', 'DurFAC')
        set_cell(123, 'HorFCA', 'DurFCA')
        set_cell(127, 'HorMAM', 'DurMAM')
        set_cell(129, 'HorMAA', 'DurMAA')
        set_cell(132, 'HorPRM', 'DurPRM')
        set_cell(133, 'HorIST', 'DurIST')
        set_cell(135, 'HorMCL', 'DurMCL')
        
        if tipomaq == "Roof Top":
            set_cell(136, 'HorMCM', 'DurMCM')
            
        if 133 in grid:
            grid[133]['trabalho'] = round(grid[133]['trabalho'] + get_val('HorMIN', 0.0), 1)
            grid[133]['duracao'] = round(grid[133]['duracao'] + get_val('DurMIN', 0.0), 1)
            
        set_cell(140, 'HorFEQ', 'DurFEQ')
        set_cell(142, 'HorLBA', 'DurLBA')
        
        if seguranca != "Não possui" and 142 in grid:
            grid[142]['trabalho'] = round(grid[142]['trabalho'] + get_val('HorMSS', 0.0), 1)
            grid[142]['duracao'] = round(grid[142]['duracao'] + get_val('DurMSS', 0.0), 1)
            
        set_cell(145, 'HorINT', 'DurINT')
        set_cell(147, 'HorTES', 'DurTES')
        set_cell(149, 'HorINS', 'DurINS')
        set_cell(150, 'HorPEE', 'DurPEE')
        set_cell(151, 'HorPEM', 'DurPEM')
        set_cell(152, 'HorFEC', 'DurFEC')
        set_cell(130, 'HorFEA', 'DurFEA')
        
        if chk_filho: delete_excel_rows(154, 154)
        if tipoestrutura == "ESSW (mecânica)": delete_excel_rows(131, 152)
        if not testesw and tipoestrutura != "ESSW (mecânica)": delete_excel_rows(148, 148)
        if not betim: delete_excel_rows(143, 143)
        if tipoestrutura != "ESSW (mecânica)": delete_excel_rows(139, 139)
        if tipoestrutura != "ESSW (mecânica)": delete_excel_rows(137, 138)
        if tipomaq != "Roof Top" and tipoestrutura != "ESSW (mecânica)": delete_excel_rows(136, 136)
        
        if nrmodulos_str in ["1 Módulo", "2 Módulos", "3 Módulos"]: delete_excel_rows(107, 118)
        if nrmodulos_str in ["1 Módulo", "2 Módulos"]: delete_excel_rows(95, 106)
        if nrmodulos_str == "1 Módulo": delete_excel_rows(83, 94)
        
        delete_excel_rows(79, 81)
        if seguranca == "Não possui" or tipoestrutura == "ESSW (mecânica)": delete_excel_rows(53, 56)
        delete_excel_rows(25, 25)
        delete_excel_rows(19, 21)
        delete_excel_rows(13, 15)
        
        if sem_eng and tipoestrutura == "Container Solar":
            delete_excel_rows(4, 57)

    # =========================================================================
    # =========================================================================
    # CENÁRIO 2: ELETROCENTRO PADRÃO
    # =========================================================================
    elif cenario_id == "eletrocentro_padrao":
        if chk_filho and 243 in grid: grid[243]['tarefa'] = 899
        if 221 in grid: grid[221]['descricao_tarefa'] = "SII - Separação Almox. Instal./ Inc."
        if 223 in grid: grid[223]['descricao_tarefa'] = "MII - Montagem Instalações / Incêndio"
        if seguranca != "Não possui" and 232 in grid:
            grid[232]['descricao_tarefa'] = "LBS - Leito e Bandejamento / Sist. Seg."

        set_cell(4, 'HorLOM', 'DurLOM')
        set_cell(8, 'HorLMM', 'DurLMM')
        if 10 in grid:
            grid[10]['trabalho'] = round(get_val('HorPBS') + get_val('HorPPA'), 1)
            grid[10]['duracao'] = round(get_val('DurPBS') + get_val('DurPPA'), 1)
            grid[10]['calculado'] = True
        set_cell(13, 'HorPCI', 'DurPCI')
        set_cell(16, 'HorPCE', 'DurPCE')
        set_cell(19, 'HorPAC', 'DurPAC')
        set_cell(21, 'HorLCA', 'DurLCA')
        set_cell(24, 'HorLAA', 'DurLAA')
        set_cell(27, 'HorLAM', 'DurLAM')
        set_cell(30, 'HorLMA', 'DurLMA')
        set_cell(34, 'HorPTR', 'DurPTR')

        set_cell(38, 'HorPIL', 'DurPIL')
        set_cell(40, 'HorLMT', 'DurLMT')
        set_cell(42, 'HorPCL', 'DurPCL')
        set_cell(44, 'HorLMC', 'DurLMC')
        set_cell(46, 'HorPIN', 'DurPIN')
        set_cell(48, 'HorLMI', 'DurLMI')
        
        if seguranca != "Não possui":
            set_cell(50, 'HorPSS', 'DurPSS')
            set_cell(52, 'HorLMS', 'DurLMS')
            
        set_cell(54, 'HorDIN', 'DurDIN')
        set_cell(56, 'HorLMD', 'DurLMD')
        set_cell(58, 'HorPBA', 'DurPBA')
        set_cell(60, 'HorLMB', 'DurLMB')
        set_cell(62, 'HorPRF', 'DurPRF')
        
        # Processos
        set_cell(11, 'Hor531', 'Dur531')
        set_cell(14, 'Hor551', 'Dur551')
        set_cell(17, 'Hor561', 'Dur561')
        set_cell(25, 'Hor581', 'Dur581')
        set_cell(28, 'Hor585', 'Dur585')
        set_cell(31, 'Hor589', 'Dur589')
        
        # Módulos 1 a 8
        modules_map = [
            (1, 66, 68, 69, 70, 71, 72, 73, 75, 78, 81),
            (2, 84, 86, 87, 88, 89, 90, 91, 93, 96, 99),
            (3, 102, 104, 105, 106, 107, 108, 109, 111, 114, 117),
            (4, 120, 122, 123, 124, 125, 126, 127, 129, 132, 135),
            (5, 138, 140, 141, 142, 143, 144, 145, 147, 150, 153),
            (6, 156, 158, 159, 160, 161, 162, 163, 165, 168, 171),
            (7, 174, 176, 177, 178, 179, 180, 181, 183, 186, 189),
            (8, 192, 194, 195, 196, 197, 198, 199, 201, 204, 207),
        ]
        for m_num, r_cor, r_fch, r_prb, r_sba, r_pre, r_ses, r_edf, r_pin, r_chi, r_che in modules_map:
            set_cell(r_cor, f'HorCOR{m_num}', f'DurCOR{m_num}')
            set_cell(r_fch, f'HorFCH{m_num}', f'DurFCH{m_num}')
            set_cell(r_prb, f'HorPRB{m_num}', f'DurPRB{m_num}')
            set_cell(r_sba, f'HorSBA{m_num}', f'DurSBA{m_num}')
            set_cell(r_pre, f'HorPRE{m_num}', f'DurPRE{m_num}')
            set_cell(r_ses, f'HorSES{m_num}', f'DurSES{m_num}')
            set_cell(r_edf, f'HorEDF{m_num}', f'DurEDF{m_num}')
            set_cell(r_pin, f'HorPIN{m_num}', f'DurPIN{m_num}')
            set_cell(r_chi, f'HorCHI{m_num}', f'DurCHI{m_num}')
            set_cell(r_che, f'HorCHE{m_num}', f'DurCHE{m_num}')
            
        set_cell(210, 'HorFAC', 'DurFAC')
        set_cell(213, 'HorFCA', 'DurFCA')
        set_cell(217, 'HorMAM', 'DurMAM')
        set_cell(219, 'HorMAA', 'DurMAA')
        set_cell(222, 'HorPRM', 'DurPRM')
        set_cell(223, 'HorIST', 'DurIST')
        set_cell(225, 'HorMCL', 'DurMCL')
        
        if tipomaq == "Roof Top":
            set_cell(226, 'HorMCM', 'DurMCM')
            
        if 223 in grid:
            grid[223]['trabalho'] = round(grid[223]['trabalho'] + get_val('HorMIN', 0.0), 1)
            grid[223]['duracao'] = round(grid[223]['duracao'] + get_val('DurMIN', 0.0), 1)
            
        set_cell(230, 'HorFEQ', 'DurFEQ')
        set_cell(232, 'HorLBA', 'DurLBA')
        
        if seguranca != "Não possui" and 232 in grid:
            grid[232]['trabalho'] = round(grid[232]['trabalho'] + get_val('HorMSS', 0.0), 1)
            grid[232]['duracao'] = round(grid[232]['duracao'] + get_val('DurMSS', 0.0), 1)
            
        set_cell(235, 'HorINT', 'DurINT')
        set_cell(237, 'HorTES', 'DurTES')
        set_cell(239, 'HorINS', 'DurINS')
        set_cell(240, 'HorPEE', 'DurPEE')
        set_cell(241, 'HorPEM', 'DurPEM')
        set_cell(242, 'HorFEC', 'DurFEC')
        set_cell(220, 'HorFEA', 'DurFEA')
        
        sorted_keys = sorted(grid.keys())
        task_list = [grid[k] for k in sorted_keys]

        def del_rows(min_r, max_r):
            nonlocal task_list
            min_i = min_r - 4
            max_i = max_r - 4
            task_list = [t for i, t in enumerate(task_list) if i < min_i or i > max_i]

        if chk_filho: del_rows(244, 244)
        if not testesw: del_rows(238, 238)
        if not betim: del_rows(233, 233)
        del_rows(229, 229)
        del_rows(227, 228)
        if tipomaq != "Roof Top": del_rows(226, 226)
        
        if nrmodulos_str in ["1 Módulo", "2 Módulos", "3 Módulos", "4 Módulos", "5 Módulos", "6 Módulos", "7 Módulos", "1"]:
            del_rows(191, 208)
        if nrmodulos_str in ["1 Módulo", "2 Módulos", "3 Módulos", "4 Módulos", "5 Módulos", "6 Módulos", "1"]:
            del_rows(173, 190)
        if nrmodulos_str in ["1 Módulo", "2 Módulos", "3 Módulos", "4 Módulos", "5 Módulos", "1"]:
            del_rows(155, 172)
        if nrmodulos_str in ["1 Módulo", "2 Módulos", "3 Módulos", "4 Módulos", "1"]:
            del_rows(137, 154)
        if nrmodulos_str in ["1 Módulo", "2 Módulos", "3 Módulos", "1"]:
            del_rows(119, 136)
        if nrmodulos_str in ["1 Módulo", "2 Módulos", "1"]:
            del_rows(101, 118)
        if nrmodulos_str in ["1 Módulo", "1"]:
            del_rows(83, 100)
            
        del_rows(82, 82)
        if tipoestrutura == "Móvel": del_rows(71, 72)
        del_rows(67, 67)
        if seguranca == "Não possui": del_rows(50, 53)
        del_rows(22, 22)
        
        if sem_eng and tipoestrutura == "Móvel":
            del_rows(4, 60)

        grid = {i + 4: t for i, t in enumerate(task_list)}

    # =========================================================================
    # CENÁRIOS 3 & 4: SKID MECÂNICA
    # =========================================================================
    elif cenario_id in ["skid_mecanica_sem_betim", "skid_mecanica_com_betim"]:
        if chk_filho:
            last_r = max(grid.keys())
            grid[last_r]['tarefa'] = 899
            delete_excel_rows(last_r, last_r)
        if sem_eng:
            delete_excel_rows(4, 15)

    # =========================================================================
    # CENÁRIO 5: ESSW ELÉTRICA
    # =========================================================================
    elif cenario_id == "essw_eletrica":
        if chk_filho:
            if 15 in grid: grid[15]['tarefa'] = 899
            delete_excel_rows(16, 16)

    # =========================================================================
    # CENÁRIO 6: PILOTIS
    # =========================================================================
    elif cenario_id == "pilotis":
        if chk_filho:
            if 8 in grid: grid[8]['tarefa'] = 899
            delete_excel_rows(9, 9)

    # =========================================================================
    # CENÁRIO 7: SKID COM ELÉTRICA (BETIM)
    # =========================================================================
    elif cenario_id == "skid_com_eletrica_betim_true":
        if chk_filho:
            if 69 in grid: grid[69]['tarefa'] = 899
            delete_excel_rows(70, 70)
        if sem_eng:
            delete_excel_rows(4, 37)

    # =========================================================================
    # CENÁRIO 8: SERVIÇO DE ENGENHARIA
    # =========================================================================
    elif cenario_id == "servico_engenharia":
        set_cell(4, 'HorLOM', 'DurLOM')
        set_cell(8, 'HorLMM', 'DurLMM')
        set_cell(9, 'HorPBS', 'DurPBS')
        set_cell(10, 'HorPCI', 'DurPCI')
        set_cell(11, 'HorPCE', 'DurPCE')
        set_cell(12, 'HorPAC', 'DurPAC')
        set_cell(14, 'HorLCA', 'DurLCA')
        set_cell(15, 'HorLAA', 'DurLAA')
        set_cell(16, 'HorLAM', 'DurLAM')
        set_cell(17, 'HorLMA', 'DurLMA')
        set_cell(19, 'HorPTR', 'DurPTR')

    # =========================================================================
    # PÓS-PROCESSAMENTO: OPERAÇÕES CONDICIONAIS & FORMATAÇÃO
    # =========================================================================
    final_tasks = []
    sorted_rows = sorted(grid.keys())
    
    for r in sorted_rows:
        item = grid[r]
        t_raw = str(item['tarefa']).strip()
        t_num = f"{int(t_raw):04d}" if t_raw.isdigit() else t_raw
        t_desc = str(item['descricao_tarefa']).strip()
        
        # Inserção da ERE antes de 0760 CHI (ou após a pintura)
        if t_num == "0760" and cenario_id == "eletrocentro_padrao":
            final_tasks.append({
                'row': 75.5,
                'tarefa': "0756",
                'tarefa_formatada': "0756",
                'descricao_tarefa': "ERE - Emissão de Relatório",
                'duracao': 1.0,
                'unidade': "DIA",
                'trabalho': 0.1,
                'calculado': False,
                'disciplina': "Pintura"
            })

        if (t_num in ["0753", "0749", "0750"]) and t_desc.startswith("PIN"):
            if betim:
                item['tarefa'] = "0749"
                item['descricao_tarefa'] = "ESU - Envio para Subcontratação - PIN"
                item['duracao'] = 1.0
                item['trabalho'] = 0.1
                item['unidade'] = "DIA"
                
            item['disciplina'] = classificar_disciplina_tarefa(item['tarefa'], item['descricao_tarefa'])
            final_tasks.append(item)
        else:
            item['disciplina'] = classificar_disciplina_tarefa(item['tarefa'], item['descricao_tarefa'])
            final_tasks.append(item)
            
    for idx, t in enumerate(final_tasks):
        t['ordem'] = idx + 1
        t_raw = str(t['tarefa']).strip()
        if t_raw.isdigit():
            t['tarefa_formatada'] = f"{int(t_raw):04d}"
        else:
            t['tarefa_formatada'] = t_raw

    total_horas = sum(t['trabalho'] for t in final_tasks)
    total_dias = sum(t['duracao'] for t in final_tasks)

    # Identificação do Seletor
    templates_definidos = {}
    if seletor_rows:
        try:
            from backend.seletor_matcher import match_seletor
            # fallback or find matching row
            pass
        except:
            pass

    return {
        "cenario_id": cenario_id,
        "cenario_descricao": cenario.get('descricao', ''),
        "qtd_tarefas": len(final_tasks),
        "total_horas": round(total_horas, 2),
        "total_dias": round(total_dias, 1),
        "tarefas": final_tasks
    }
