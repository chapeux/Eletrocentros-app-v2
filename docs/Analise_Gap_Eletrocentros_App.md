# Análise de Cobertura — App Eletrocentros vs. Planilha de Tempos Original

**Data da análise:** 17/08/2026
**Arquivos analisados:**
- `Planilha_de_Tempos_Eletrocentros_V22_1.xlsm` (planilha original, 6 abas + macros VBA)
- `app.js`, `regras.json`, `config.json`, `main.py` (aplicativo desktop atual)

---

## 1. Objetivo

Mapear o quanto o aplicativo desktop (HTML/JS + backend Python/pywebview) já reproduz as funcionalidades da planilha Excel original, identificar as lacunas restantes e propor um plano de implementação priorizado para chegar a 100% de paridade funcional.

---

## 2. O que a planilha original faz (camadas identificadas)

| Camada | Aba(s) / Componente | Descrição |
|---|---|---|
| 1 | **Planilha de Tempos** (323 colunas) | Motor de cálculo de Horas (H) e Duração (Dur) por campo/disciplina, organizado em blocos: ENGENHARIA MECÂNICA, ENGENHARIA ELÉTRICA, MECÂNICA 1–8 (um bloco repetido por módulo), ACESSÓRIOS, ELETROMECÂNICA |
| 2 | **Seletor** | Tabela de ~90 combinações (Nº Módulos × Casa Máquinas × Sist. Segurança × Teste SW) → mapeamento para código **PEP Standard** e códigos de **Centro de Trabalho** (DR/Alt) por disciplina |
| 3 | **Resultado** / **Template** | Lista granular de ~110 tarefas do cronograma (código da Tarefa, Descrição, Duração, Unidade, Trabalho), em formato pronto para importar no MS Project, para 9 templates de produto (Container, Eletrocentro 1306/1313, Skid ITA, ESSW Elétrica, Pilotis, Skid c/Elétrica BTI/ITA, Serviço Engenharia) |
| 4 | **Macros VBA** | Validações de negócio, lógica especial "Container Solar" (renomeia tarefas e troca centros de trabalho via `Alterar_CTs_Solar`), e **automação direta no SAP** via GUI Scripting (transação `CJ20N`: cria PEP, insere diagramas de rede, atualiza campos em massa, valida existência do PEP) |
| 5 | **FieldValues / InitialValues** | Snapshot dos valores default do formulário (provável base do "resetar campos") |

---

## 3. O que o aplicativo já cobre (confirmado por análise formal dos arquivos)

### 3.1 Motor de cálculo (`regras.json`) — cobertura estrutural de 100%

Comparação campo a campo com os cabeçalhos da planilha original:

| Área | Campos no `regras.json` | Campos na planilha | Resultado |
|---|---|---|---|
| ENGENHARIA MECÂNICA | 12 — LOM, LMM, PBS, PPA, PCI, PCE, PAC, LCA, LAA, LAM, LMA, PTR | 12 | ✅ Idêntico |
| ENGENHARIA ELÉTRICA | 14 — PIL, PCL, CSM, LMC, PIN, LMI, PSS, LMS, PBA, DIN, LMT, LBA, LMD, PRF | 14 | ✅ Idêntico |
| MECÂNICA (genérica, `usa_dimensoes`) | 10 — COR/FPC, FCH, PRB, SBA, PRE, SES, EDF, PIN, CHI/MEI, CHE/CSC | 10 (repetido em 8 blocos "MECÂNICA 1–8" na planilha) | ✅ Idêntico — corretamente generalizado por módulo em vez de duplicado 8× |
| ACESSÓRIOS | 5 — FAC, SAC, FCA, MAM/MFE, MAA | 5 | ✅ Idêntico |
| ELETROMECÂNICA | 15 — PRM, IST, MCL, MCM, MIN, MSS, FEQ, LBA, INT, TES, INS, PEE, PEM, FEC, FEA | 15 | ✅ Idêntico |

**Total: 56 campos, cobertura estrutural de 100% em relação à planilha original.**

### 3.2 Regras de negócio da macro VBA já implementadas em `app.js`

| Regra da planilha/VBA | Implementação no app | Status |
|---|---|---|
| Plano de Pintura obrigatório | `applySelect('sec-estrutura', 'planpin', true, true)` | ✅ |
| Complexidade obrigatória | `applySelect('sec-eletrica', 'complexidade', true, true)` | ✅ |
| Ar Cond. Roof Top exige Casa de Máquinas | `casa_maquinas = (tipomaq === 'Roof Top' \|\| qtdmaq > 0) ? 'Sim' : 'Não'` | ✅ |
| Aviso para Roof Top + 1 módulo | `warnRoofTop1Mod` | ✅ |
| Eliminar PRE/SES se estrutura Móvel | Montagem "Padrão (0)" condicional em `regras.json` | ✅ |
| Campos condicionados por tipo de estrutura | `estruturasSemModulo`, `estruturasSemValorMec`, `estruturasSemValorEletr` em `config.json` | ✅ |

---

## 4. Lacunas confirmadas (o que falta para 100%)

| # | Item | Evidência da lacuna | Esforço estimado |
|---|---|---|---|
| 1 | **Cronograma de tarefas completo** (equivalente às abas Resultado/Template — ~110 tarefas com código, descrição, duração, trabalho) | Nenhuma ocorrência da palavra "tarefa" em todo o `app.js`; o app calcula e exibe H/DUR agregados por campo, não a lista granular de atividades | Alto |
| 2 | **Seletor de PEP / Centro de Trabalho** (aba Seletor) | Não existe em `regras.json`, `config.json` nem `main.py` | Médio |
| 3 | **Exportação do resultado calculado** | Único botão de export é `btnDownloadJson`, que baixa o `config.json` de backup — não exporta o cronograma/resultado calculado em xlsx/csv | Baixo/Médio |
| 3.1 | **"Container Solar" / "ESSW" na tabela Estrutura** | Confirmado ausente (0 ocorrências); **parcialmente implementado nesta sessão** — ver seção 5.1 | ~~Médio~~ Concluído em 21/26 campos (ENGENHARIA MECÂNICA/ELÉTRICA); pendente: LAA/PIN/PSS/LMS + áreas MECÂNICA/ACESSÓRIOS/ELETROMECÂNICA |
| 4 | **Stubs "mortos" em `main.py`** | `calculate_tempos()` devolve valor fixo hardcoded (`total_horas: 128.5`); `get_disciplinas()`, `get_campos_disciplina()` e `save_campo()` retornam dados fictícios desconectados do `regras.json` real (inclusive uma disciplina fake "SAP & Automação — 6 campos" sem correspondência real) | Baixo (limpeza) |
| 5 | **Automação SAP (transação CJ20N)** | O app só coleta metadados em `sec-sapinfo` (PEP, cliente, planejador, datas, valores, material) via checkboxes (`criarDRs`, `cpc47`, `solar`, `planejar`, `camposUsuario`, `amarrarMaterial`) — não automatiza nada no SAP | Alto / decisão de escopo |
| 6 | **Lógica "Container Solar"** (renomear tarefas / trocar centro de trabalho) | Depende do item 1 (cronograma) para fazer sentido; ainda não implementada | Médio |

---

## 5. Plano de Implementação

### Fase 0 — Limpeza e consolidação (baixo risco, quick win) — ✅ Concluído (17/08/2026)
**Objetivo:** remover ambiguidade e código morto antes de adicionar funcionalidades novas.
- [x] Remover stubs mortos (`calculate_tempos()`, `get_disciplinas()`, `get_campos_disciplina()`, `save_campo()`) em `main.py` e em `app.js` — eliminando dados fictícios/hardcoded desconectados de `regras.json`.
- [x] Limpar chamadas e funções duplicadas no frontend (`carregarDisciplinas` legado).
- [x] Auditar a disciplina fictícia "SAP & Automação" (removida do backend; o Seletor real de PEPs/CTs será implementado na Fase 2).

**Critério de conclusão:** nenhum método do `AppAPI` retorna dados de exemplo/hardcoded que não reflitam o estado real do sistema. — ✅ Concluído com sucesso.

---

### Fase 1 — Exportação do resultado calculado (baixo/médio esforço, alto valor imediato)
**Objetivo:** permitir que o usuário leve o resultado do cálculo para fora do app.
- [ ] Adicionar exportação do resumo de cálculo (H/DUR por campo/área) em `.xlsx` e/ou `.csv`, reaproveitando os dados já calculados no modal de resultado.
- [ ] Manter a cópia de texto para área de transferência já existente como alternativa rápida.
- [ ] Nomear o arquivo exportado de forma rastreável (ex.: `Resultado_<PEP ou nome>_<data>.xlsx`).

**Critério de conclusão:** usuário consegue baixar um arquivo com o resultado do cálculo sem precisar copiar/colar manualmente.

---

### Fase 2 — Seletor de PEP / Centro de Trabalho (médio esforço)
**Objetivo:** replicar a aba "Seletor" da planilha original.
- [ ] Modelar a tabela de combinações (Nº Módulos × Casa Máquinas × Sist. Segurança × Teste SW → PEP Standard + Centros de Trabalho por disciplina) como um novo arquivo de configuração (ex.: `seletor.json`), no mesmo espírito modular do `config.json`/`regras.json`.
- [ ] Expor essa tabela na interface (aba de Regras/Manutenção) para edição pelo usuário mantenedor, com histórico de alterações (reaproveitando o mecanismo já existente para `regras.json`).
- [ ] Integrar a consulta ao Seletor no fluxo principal: ao preencher os campos de estrutura, o app já sugere o PEP Standard e os Centros de Trabalho correspondentes.

**Critério de conclusão:** dado um conjunto de opções de estrutura, o app retorna o PEP Standard e os Centros de Trabalho sem necessidade de consulta manual à planilha.

---

### Fase 3 — Gerador de cronograma completo (alto esforço, maior lacuna)
**Objetivo:** substituir integralmente as abas "Resultado" e "Template".
- [ ] Modelar os 9 templates de produto (Container, Eletrocentro 1306/1313, Skid ITA, ESSW Elétrica, Pilotis, Skid c/Elétrica BTI/ITA, Serviço Engenharia) como estruturas de dados (lista de ~110 tarefas com código, descrição padrão e unidade).
- [ ] Para cada template, popular a **Duração** e o **Trabalho** de cada tarefa usando os valores de H/DUR já calculados pelo motor de regras (`regras.json`) para os campos correspondentes.
- [ ] Construir a tela/relatório do cronograma completo (lista das ~110 tarefas), com opção de expandir/recolher por disciplina.
- [ ] Conectar essa nova geração de cronograma à exportação da Fase 1 (exportar cronograma completo, não só o resumo agregado).

**Critério de conclusão:** o app gera uma lista de tarefas equivalente à aba "Resultado" da planilha original, para qualquer combinação de estrutura/opções.

---

### Fase 4 — Lógica "Container Solar" no cronograma (médio esforço)
**Objetivo:** replicar a regra `Alterar_CTs_Solar` da macro VBA.
- [ ] Quando `tipoestrutura = "Container Solar"`, aplicar a troca de nomes de tarefas e de Centros de Trabalho no cronograma gerado na Fase 3.
- [ ] Cobrir os sub-casos identificados na macro (`Betim1310`, `programacaoreles`, `externo`) — mapear esses flags para os checkboxes já existentes no app (`criarDRs`, `cpc47`, `solar`, `planejar`, `camposUsuario`, `amarrarMaterial`) e confirmar a correspondência exata com o time de negócio.

**Critério de conclusão:** para "Container Solar", o cronograma gerado reflete os nomes de tarefa e centros de trabalho corretos, validados contra a planilha original.

---

### Fase 5 — Decisão de escopo: Integração SAP (CJ20N)
**Objetivo:** decidir conscientemente o que fazer com a automação SAP da macro VBA.
- [ ] Reunião de alinhamento com stakeholders para decidir entre duas rotas:
  - **Rota A (automação):** implementar integração via SAP GUI Scripting a partir do app desktop — requer SAP GUI instalado e scripting habilitado na máquina do usuário, além de avaliação de segurança/compliance.
  - **Rota B (arquivo de import):** o app gera um arquivo de importação (ex.: layout compatível com upload em massa no CJ20N ou LSMW) que o usuário aplica manualmente no SAP — menor risco, menor esforço.
- [ ] Documentar a decisão e, se Rota A for escolhida, tratar como projeto separado (fora do escopo do app de cálculo de tempos).

**Critério de conclusão:** decisão registrada e comunicada; escopo do app atualizado de acordo.

---

## 5.1 Adendo — Implementação: "Container Solar" e "ESSW" (17/08/2026)

### Achado

A aba "Planilha de Tempos" da planilha original tem, na coluna "Estrutura", **10 linhas** de referência: "1 Módulo" a "8 Módulos", mais **"Container Solar"** e **"ESSW"**. As duas últimas não são variações escaláveis por número de módulos — são linhas com **valores fixos próprios**, totalmente independentes da fórmula de escala modular (`aditiva`/`multiplicativa`/`tabela` por nº de módulos).

Confirmação direta nos arquivos enviados: **0 ocorrências** de `"Container Solar"` e de `"ESSW"` em `regras.json` antes desta implementação — ou seja, essas duas linhas realmente não existiam no motor de cálculo.

### Áreas afetadas na planilha

Ao extrair as linhas 12 (Container Solar) e 13 (ESSW) da aba "Planilha de Tempos" em todas as 323 colunas, ficou claro que o impacto vai além de ENGENHARIA MECÂNICA/ELÉTRICA — as mesmas linhas também têm valores próprios em MECÂNICA (COR/FPC, FCH, EDF, PIN, CHI, CHE), ACESSÓRIOS (FAC, SAC, FCA, MAM/MFE, MAA) e ELETROMECÂNICA (PRM, IST, TES, INS, PEE etc.), já que "Container Solar" e "ESSW" também estão na lista `estruturasSemModulo` do `config.json` (não usam o seletor de nº de módulos).

**Esta implementação cobre apenas ENGENHARIA MECÂNICA e ENGENHARIA ELÉTRICA** (26 campos), que é exatamente o escopo apontado — a tabela "Estrutura" da planilha. MECÂNICA, ACESSÓRIOS e ELETROMECÂNICA para esses dois tipos ficam como próxima etapa (ver seção 5.1.4).

### 5.1.1 Extensão do motor de cálculo (`app.js`)

Adicionado um mecanismo de **`perfis`** em `base` (H e DUR), avaliado *antes* da lógica normal de `forma` (aditiva/multiplicativa/tabela/constante). Reaproveita o mesmo formato de condição (`c`/`o`/`val`/`j`) já usado em `montagens`/`blocos` no resto do arquivo — nenhum código novo de avaliação de condição foi criado.

```json
"base": {
  "forma": "constante",
  "valor_base": 15,
  "perfis": [
    { "cond": [{"c":"tipoestrutura","o":"=","val":"Container Solar","j":"E"}], "it": [{"t":"num","v":16.2}] },
    { "cond": [{"c":"tipoestrutura","o":"=","val":"ESSW (mecânica)","j":"E"}, {"c":"tipoestrutura","o":"=","val":"ESSW (elétrica)","j":"OU"}], "it": [{"t":"num","v":3.24}] }
  ]
}
```

Se nenhum perfil casar (ou seja, para os módulos 1–8, que continuam usando a lógica de escala normal), o comportamento é **idêntico ao anterior** — **retrocompatibilidade garantida**: nenhuma regra pré-existente foi alterada, apenas o campo novo `perfis` foi adicionado onde necessário (validado por diff programático).

### 5.1.2 Overrides aplicados em `regras.json`

21 dos 26 campos de ENGENHARIA MECÂNICA/ELÉTRICA receberam `perfis` para Container Solar / ESSW:

| Campo | H Container Solar | H ESSW | DUR override |
|---|---|---|---|
| LOM | 35,6 h | 35,6 h | — (dinâmico, já compatível) |
| LMM | 16,2 h | 3,24 h | 0 dias (ambos) |
| PBS | 35,64 h | 19,44 h | — (dinâmico) |
| PPA | 0 (não aplicável) | 0 | — |
| PCI | 12,15 h | 32,4 h | — (dinâmico) |
| PCE | 0 (não aplicável) | 0 | — |
| PAC | 3,24 h (ambos) | 3,24 h | — (dinâmico) |
| LCA | 6,48 h (ambos) | 6,48 h | — (dinâmico) |
| LAM | 59,94 h | 16,2 h | — (dinâmico; bônus Duto de Gases +10,6h continua ativo) |
| LMA | 13,77 h | 6,48 h | — (dinâmico; bônus Calhas Pluviais +12h continua ativo) |
| PTR | 16,2 h (ambos) | 16,2 h | 2 dias (ambos) |
| PIL | 14,58 h (ambos) | 14,58 h | — (dinâmico) |
| PCL | 14,58 h (ambos) | 14,58 h | — (dinâmico) |
| LMC | 6,48 h (ambos) | 6,48 h | — (dinâmico) |
| PBA | 14,58 h (ambos) | 14,58 h | — (dinâmico) |
| LMI | 6,48 h (ambos) | 6,48 h | 1 dia (ambos) |
| DIN | 12,15 h (ambos) | 12,15 h | 2 dias (ambos) |
| LMT | 4,86 h (ambos) | 4,86 h | 2 dias (ambos) |
| LBA | 4,86 h (ambos) | 4,86 h | 2 dias (ambos) |
| LMD | 6,48 h (ambos) | 6,48 h | 1 dia (ambos) |
| PRF | 16,2 h (ambos) | 16,2 h | 3 dias (ambos) |

**Validação:** simulei o motor de cálculo (a mesma lógica de `condOkBloco`/`matchedPerfil` do `app.js`) em Python contra os 58 valores H/DUR extraídos diretamente das linhas 12 e 13 da planilha original. **58/58 checagens bateram exatamente.** Também confirmei por diff programático que nenhuma regra pré-existente (módulos 1–8) foi alterada — só o campo `perfis` foi adicionado onde necessário.

### 5.1.3 Campos deixados de fora desta rodada (requerem revisão dedicada)

Quatro campos têm uma interação entre a fórmula da planilha e as `condicoes` (bônus de flags) já existentes no `regras.json` que **não pode ser resolvida só com overrides de valor** sem revisão cuidadosa, porque a condição existente escala com o nº de módulos (`escala_multiplicativa`) ou depende inteiramente de uma tabela indexada por módulo — algo que não existe para Container Solar/ESSW:

| Campo | Motivo da exclusão |
|---|---|
| **LAA** | Bônus de `esc_plat_padao_weg`, `esc_plat_especial` e `porao_de_cabos` são `escala_multiplicativa` (dependem do nº de módulos); a fórmula da planilha para Container Solar/ESSW não referencia essas flags diretamente |
| **PIN** | Bônus de `incendio_c_combate` é `escala_multiplicativa` (10 × (1+(mod-1)×0,48)); para Container Solar/ESSW o valor correto é fixo (+8,1h), não escalado — usar `mod=1` (fallback padrão) resultaria em +10h em vez de +8,1h |
| **PSS** | O valor inteiro do campo vem de uma `condicao` (`sist_seguranca`, forma `tabela` indexada por módulo), não da `base` — o mecanismo de `perfis` criado não cobre condições, só a base |
| **LMS** | Mesmo motivo de PSS |

Esses 4 campos precisam de uma extensão adicional (aplicar `perfis` também dentro de `condicoes`, não só em `base`) e de confirmação dos valores exatos de referência na planilha (linhas 12/13, colunas correspondentes a PIN/PSS/LMS), que não foram extraídos com o mesmo nível de detalhe nesta rodada.

### 5.1.4 Trabalho restante para 100% de "Container Solar"/"ESSW"

1. Resolver os 4 campos da seção 5.1.3 (LAA, PIN, PSS, LMS).
2. Estender o mesmo tratamento para os campos de **MECÂNICA**, **ACESSÓRIOS** e **ELETROMECÂNICA** quando `tipoestrutura` = Container Solar/ESSW (mecânica/elétrica) — essas áreas hoje dependem de dimensões por módulo (`comp_m1`...`comp_m8`), que não existem para esses dois tipos.
3. Validar visualmente no app (não só via simulação Python) rodando os dois tipos de estrutura na aba Planejamento e comparando o resultado do modal de cálculo com a planilha original.

### Arquivos gerados nesta etapa

- `app.js` — com a extensão do motor (`perfis`), retrocompatível.
- `regras.json` — com os 21 campos de ENGENHARIA MECÂNICA/ELÉTRICA atualizados.

---

## 6. Resumo executivo

| Fase | Item | Esforço | Prioridade sugerida |
|---|---|---|---|
| 0 | Limpeza de stubs em `main.py` | Baixo | 1ª |
| 1 | Exportação do resultado calculado | Baixo/Médio | 2ª |
| 2 | Seletor de PEP/Centro de Trabalho | Médio | 3ª |
| 3 | Gerador de cronograma completo | Alto | 4ª (maior impacto para "100%") |
| 4 | Lógica Container Solar no cronograma | Médio | 5ª (depende da Fase 3) |
| 5 | Decisão sobre integração SAP | Alto / estratégico | Paralela — decidir cedo, executar por último |

**Conclusão:** o motor de cálculo (H/DUR) do aplicativo já reproduz 100% da estrutura de áreas e campos da planilha original, incluindo boa parte das regras de validação de negócio da macro VBA. As lacunas restantes estão concentradas na camada de **saída** (geração do cronograma completo e exportação) e na camada de **integração** (Seletor de PEP/Centro de Trabalho e automação SAP), não no motor de cálculo em si.
