# Contexto e Plano do Projeto — Eletrocentros App v2

## Resumo do Projeto
Interface web desktop unificada e 100% modular (HTML5/CSS3/JS) para substituição de GUIs antigas em Tkinter por uma aplicação moderna conectada a um backend Python (`main.py`) e alimentada por arquivos de configurações externos (`config.json` e `regras.json`).

## Arquitetura e Componentes
1. **Configuração Externa (`config.json` & `regras.json`)**:
   - `config.json`: Armazena todas as listas de seletores (Estrutura, Pintura, Ar Condicionado, Incêndio, Segurança, Complexidade, Planejadores) e regras de aplicação.
   - `regras.json`: Estruturado em 3 níveis principais:
     1. `area`: Categoria principal da disciplina (ex: "ENGENHARIA MECÂNICA", "ENGENHARIA ELÉTRICA").
     2. `campos`: Mapeamento de códigos (LOM, LMM, PBS, PPA, PCI, PCE, PAC, LCA, LAA, LAM, LMA, PTR, etc.).
     3. Sub-abas `H` (Horas) e `DUR` (Duração): Cada sub-aba possui sua regra de cálculo base (constante, aditiva, multiplicativa, tabela ou derivado_h com pipeline de etapas) e sua lista de condições adicionais.
   - Suporte ao Pipeline de Etapas Operacionais em `derivado_h` (`base.etapas` array): Permite criar, reordenar, alterar e remover passos sequenciais de cálculo sobre o valor de Horas (H) (Dividir ÷, Multiplicar ×, Somar +, Subtrair -, Arredondar ⌈⌉/⌊⌋).
   - Permite edição através da interface visual amigável ou diretamente via editor RAW JSON (`Regras JSON` no modal de configurações).

2. **Ativos & Ícones (`assets/`)**:
   - Pasta dedicada contendo `icone.ico` para personalização do ícone da aplicação.
   - Ícone aplicado via Win32 API (`WM_SETICON` para ícones de canto de janela e barra de tarefas do Windows) no launcher Python `main.py` e favicon HTML no browser.

3. **Frontend (`frontend/`)**:
   - `index.html`: Interface unificada nativa com abas de **Planejamento PCP** e **Regras**.
   - `style.css`: Design System com suporte a temas Dark/Light, scrollbars customizados escuros, sem nenhuma sombra projetada ou sombra interna, com caixa retangular padronizada no editor JSON (`border-radius: 9px`), scroll oculto na lista de campos (`.col-list`) e scroll visível no painel de edição (`.col-editor`).
   - `app.js`: Carregamento dinâmico do `config.json` e `regras.json` via API Python/fetch, gerenciamento de abas, modal de credenciais (`admin`/`1234`), sub-abas H / DUR por campo, suporte a fórmulas derivadas (`derivado_h` com pipeline de passos reordenáveis `base.etapas`), operações matemáticas adicionais, editor visual de listas e editor de RAW JSON para `regras.json`.

4. **Backend Python (`main.py`)**:
   - Ponto de entrada Desktop utilizando `pywebview` (janela nativa Chromium/Edge WebView2) com fallback para servidor HTTP local (`http.server`).
   - Rotina `apply_native_window_icon()` que injeta o `icone.ico` nas janenas nativas do Windows (`WM_SETICON`).
   - Classe `AppAPI` expondo os métodos:
     - `get_config()` & `save_config(data)` (leitura e gravação do `config.json` no disco).
     - `get_regras()` & `save_regras(data)` (leitura e gravação do `regras.json` no disco).
     - `verify_password(password)` (autenticação de mantenedor).
     - `calculate_tempos()`.

## Registro de Alterações Solicitadas pelo Usuário
- [x] Unificação dos protótipos em `docs/` (`dados_planejamento.html` e `painel_manutencao.html`).
- [x] Modal de credenciais para a aba do Painel de Regras.
- [x] Backend Python `main.py` com launcher desktop sem Tkinter.
- [x] Compactar o contexto completo do projeto neste arquivo `plan.md`.
- [x] Ajustar scrollbar clara no Tema Escuro para scrollbar escura combinando com o tema.
- [x] Remover rolagem externa duplicada na aba de Planejamento.
- [x] Remover a moldura/barra de título simulada (estética fake) para que o aplicativo ocupe 100% da janela nativa.
- [x] Remover sombras internas (`box-shadow: inset`) e linhas de brilho nos painéis.
- [x] Eliminar vazamentos de sombras laterais do painel de histórico (`hist-panel`), modais, seletores e foco.
- [x] Tornar a aplicação 100% modular com arquivo de configurações externo (`config.json`), permitindo edições pelos usuários sem alterar o código.
- [x] Adicionar seletor amigável no menu de configurações para escolha de lista a editar (Restrito a usuários autenticados/modo mantenedor).
- [x] Corrigir formato do editor JSON no modal de configurações de 99px (formato pílula/oval) para caixa retangular padrão (`border-radius: 9px`).
- [x] Criar pasta `assets/` e alocar o arquivo `icone.ico`.
- [x] Aplicar rotina Win32 API (`WM_SETICON` ICON_SMALL e ICON_BIG) para substituir o ícone padrão do Python pelo `icone.ico` tanto na barra de tarefas do Windows quanto no canto superior esquerdo da barra de título da janela.
- [x] Renomear a segunda aba da barra de navegação superior de **Manutenção** para **Regras**.
- [x] Reestruturar a página de Regras para ler o arquivo `regras.json` com primeiro nível `area`, campos (LOM, LMM, PBS, PPA, PCI, PCE, PAC, LCA, LAA, LAM, LMA, PTR...), e sub-abas `H` e `DUR` com regras e condições próprias.
- [x] Configurar scrollbar oculta (funcional com rolagem por roda do mouse, mas sem barra visível) na coluna de campos (`.col-list`) e scrollbar visível no painel de edição de condições (`.col-editor`).
- [x] Adicionar a aba e editor RAW JSON de **Regras JSON** no modal de configurações para permitir a edição direta e salvamento persistente do arquivo `regras.json`.
- [x] Corrigir a ativação do botão de salvar ao adicionar, remover ou alterar condições adicionais nas regras.
- [x] Criar o tipo de cálculo base `derivado_h` (Fórmula Operacional) para a sub-aba DUR, calculando `ROUNDUP(H / 7.92) - 1` de forma 100% modular e configurável via UI e JSON.
- [x] Ajustar alinhamento horizontal dos 3 campos do `derivado_h` (Dividir H por, Arredondamento, Subtrair no final) e corrigir a resolução e passagem de `campoObj` para extrair o valor de H e testar todas as opções/condições de H no preview de DUR.
- [x] Implementar o **Pipeline Modular de Etapas Sequenciais (`base.etapas`)**, permitindo ao usuário adicionar, remover, trocar (Dividir ÷, Multiplicar ×, Somar +, Subtrair -, Arredondar ⌈⌉/⌊⌋) e reordenar (↑/↓) qualquer quantidade de passos de cálculo na fórmula de horas H.
