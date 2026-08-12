# Contexto — Regras MECÂNICA com construtor de blocos condicionais

> Substitui o `contexto-regras-mecanica.md` anterior. A abordagem de "forma: componentes"
> (tabela de fatores + matriz) foi **descartada** em favor de um construtor de blocos genérico.

## Projeto

App desktop (pywebview + `main.py`), frontend `frontend/index.html` / `app.js` / `style.css`.
Configuração externa em `config.json` (listas + regras de aplicação) e `regras.json`
(áreas → campos → sub-abas `H` e `DUR`).

Áreas existentes: `ENGENHARIA MECÂNICA`, `ENGENHARIA ELÉTRICA` — calculam só por quantidade de
módulos, sem dimensões. Motor atual: `valorBase(base, mod, campoObj, flagsAtivos, vH)`, formas
`constante`, `multiplicativa`, `degrau_fixo`, `aditiva`, `tabela`, `derivado_h`.

## Áreas novas a criar

| Área | Depende de dimensões? | Esforço |
|---|---|---|
| **MECÂNICA** (MEC1–MEC8) | Sim (comp × larg por módulo) | Nova forma `blocos` no motor |
| **ACESSÓRIOS** | Não | Só JSON, reaproveita formas existentes |
| **ELETROMECÂNICA** | Não | Só JSON, reaproveita formas existentes |

## O problema que originou o construtor

Fórmula original do Excel (MEC1 – Corte), com três `SE` aninhados por tipo de estrutura:

```
=ARRED((SE(BT4="Móvel";((BU4*$CA$27+BV4*$CB$27)*$CC$27)/60;
  SE(BT4="Embarcado";((($CA$23*BU4+4)*$CC$23)+(2*(BU4*$CA$24+4)*$CC$24)
   +((BU4*$CA$25+BV4*$CB$25)*$CC$25)+((BU4*$CA$27+BV4*$CB$27)*$CC$27))/60;
  ((($CA$23*BU4+4)*$CC$23)+(2*(BU4*$CA$24+4)*$CC$24)+((BU4*$CA$25+BV4*$CB$25)*$CC$25)
   +(2*(BU4*$CA$26+BV4*$CB$26)*$CC$26)+((BU4*$CA$27+BV4*$CB$27)*$CC$27))/60)))*0,9*1,1;1)
```

Tabela de fatores (Excel CA23:CC27):

| Componente | Fator1 | Fator2 | MinB |
|---|---|---|---|
| Lateral | 1,14 | — | 2,815 |
| FrenFund | 1,31 | — | 1,472 |
| Teto | 1,29 | 1,19 | 4,785 |
| Telhado | 1,38 | 1,025 | 1,36 |
| Base | 1,40 | 1,405 | 14,17 |

`BU4` = comprimento, `BV4` = largura. Pós-processamento comum a todos os ramos:
`/60 × 0,9 × 1,1`, arredondado a 1 casa.

## Arquitetura escolhida

### 1. Variáveis — nova área em `config.json`

Editável pelo menu de configurações (ícone de engrenagem). Alimenta todas as listas de
variáveis do construtor.

```json
"variaveis": [
  { "grupo": "Dimensões",       "nome": "Comprimento",    "chave": "comp",   "tipo": "entrada" },
  { "grupo": "Dimensões",       "nome": "Largura",        "chave": "larg",   "tipo": "entrada" },
  { "grupo": "Dimensões",       "nome": "Altura",         "chave": "alt",    "tipo": "entrada" },
  { "grupo": "Sistema",         "nome": "Nº de módulos",  "chave": "nmod",   "tipo": "entrada" },
  { "grupo": "Tempos de Corte", "nome": "Lateral-Fator1", "chave": "lat_f1", "tipo": "constante", "valor": 1.14 },
  { "grupo": "Tempos de Corte", "nome": "Lateral-MinB",   "chave": "lat_mb", "tipo": "constante", "valor": 2.815 }
]
```

- `tipo: "entrada"` → valor vem do formulário preenchido pelo usuário.
- `tipo: "constante"` → valor fixo definido na própria tela de configuração.
- Agrupadas por `grupo` (vira `<optgroup>` nos selects).

### 2. Blocos

Um bloco é uma sequência nomeada de itens. **Cada bloco funciona como um parêntese** — é
avaliado inteiro antes de entrar na montagem, o que elimina a necessidade de parênteses avulsos
na UI.

Tipos de item: `var` (variável), `op` (operação), `num` (valor fixo), `blk` (outro bloco).
Operações disponíveis: `+ − × ÷ % ^`.

```json
"blocos": [
  { "id": "b1", "nome": "Lateral",
    "itens": [
      {"t":"var","v":"lat_f1"}, {"t":"op","v":"*"}, {"t":"var","v":"comp"},
      {"t":"op","v":"+"}, {"t":"num","v":4}
    ] }
]
```

Na UI os blocos ficam **ocultos** — só uma barra resumo ("5 blocos definidos: Lateral · Frente/Fundo
· …") com botão "Ver / criar blocos" que abre um modal. Dentro do modal, cada bloco mostra seu
resultado calculado e a linha "usado em: …" listando quais montagens o referenciam.

### 3. Montagens condicionais

Substituem os `SE` aninhados. Lista ordenada; **vale a primeira cuja condição for atendida**.
A última é sempre a padrão (`"padrao": true`), não pode ser apagada nem movida — garante que
sempre exista resultado.

```json
"montagens": [
  { "id": "m1", "nome": "Móvel",
    "condicoes": [ {"campo":"tipoestrutura","op":"=","valor":"Móvel","juncao":"E"} ],
    "itens": [ {"t":"blk","v":"b5"}, {"t":"op","v":"*"}, {"t":"var","v":"bas_mb"} ] },
  { "id": "m2", "nome": "Embarcado",
    "condicoes": [ {"campo":"tipoestrutura","op":"=","valor":"Embarcado","juncao":"E"} ],
    "itens": [ ... ] },
  { "id": "m0", "nome": "Demais estruturas", "padrao": true, "condicoes": [], "itens": [ ... ] }
]
```

- Campos de condição: `tipoestrutura`, `nmod`, `tipomaq`, `incendio` (vindos de `config.json`).
- Operadores: `= ≠ > < ≥ ≤`.
- Múltiplas condições por montagem, unidas por `E` / `OU`.
- Reordenação por setas ▲▼.

### 4. Ajuste final (pós-processamento)

**Compartilhado por todas as montagens** — reproduz o Excel, onde `/60 × 0,9 × 1,1` está fora
dos `SE`. Fica separado da cadeia da montagem de propósito: se fosse mais um item na sequência,
a precedência matemática faria o `÷60` dividir só a última parcela em vez do total.

```json
"ajuste_final": [ {"t":"op","v":"/"},{"t":"num","v":60},
                  {"t":"op","v":"*"},{"t":"num","v":0.9},
                  {"t":"op","v":"*"},{"t":"num","v":1.1} ]
```

### 5. Simulação

No rodapé da tela. Filtros: **tipo de estrutura, comprimento, largura** apenas. Mostra somente o
resultado final em horas + qual montagem foi aplicada. Sem detalhamento por bloco.

## Decisões de UI

- **Sem cores por tipo de item.** Todos os chips (variável, operação, valor fixo, bloco) usam o
  mesmo tom neutro — a distinção vem do conteúdo. A única cor funcional que resta é o verde do
  badge `APLICADA` na montagem que bate com o cenário simulado.
- Blocos ocultos atrás de botão; montagens visíveis na tela principal.
- Ordem das montagens explícita (numeração 1, 2, 3 + setas).

## Implementação no motor (`app.js`)

1. **`getVars()`** — resolve valores: `entrada` lê do formulário/DOM (`.mod-comp`, `.mod-larg`,
   `nrmodulos`, etc.), `constante` lê de `CONFIG.variaveis`.
2. **`chainExpr(itens)`** — serializa a sequência em expressão JS; item `blk` expande para
   `'(' + chainExpr(bloco.itens) + ')'` (recursivo — cuidar de referência circular).
3. **`evalChain(itens)`** — `Function('"use strict";return (' + expr + ')')()` com try/catch,
   retorna `NaN` em erro.
4. **`condOk(montagem)`** — avalia condições contra o contexto do formulário, respeitando `E`/`OU`;
   comparação numérica quando o campo é numérico.
5. **`matched()`** — primeira montagem com `condOk` verdadeiro; senão a `padrao`.
6. **`valorBase`** — novo ramo `forma === 'blocos'`: `evalChain(matched().itens.concat(ajuste_final))`.
7. **Agregação por módulo**: quando houver mais de um módulo com dimensões distintas, calcular por
   módulo e somar (`por_modulo_somado`) — cada módulo tem suas próprias laterais, teto e base.
   **A definir se confirma.**

## Pendências

- [ ] Confirmar agregação com múltiplos módulos (somar por módulo vs. comprimento total).
- [ ] **FrenFund usa comprimento no Excel** (`BU4`), não largura. Painéis de frente/fundo
      normalmente escalariam com a largura. Pode ser bug herdado — validar com o domínio.
- [ ] Decidir sobre campos de condição sem filtro na simulação (`nmod`, `tipomaq`, `incendio`):
      ou limitar a lista de campos de condição, ou fazer a barra de simulação exibir
      automaticamente os campos usados em alguma condição ativa.
- [ ] Detectar e bloquear referência circular entre blocos.
- [ ] Popular MEC1–MEC8, ACESSÓRIOS e ELETROMECÂNICA com os valores reais.
- [ ] Migrar as demais fórmulas do Excel para blocos + montagens.

## Mockups gerados (referência visual)

| Arquivo | Abordagem |
|---|---|
| `mockup-regras-mecanica.html` | v1 — tabela de fatores + matriz (descartada: densa demais) |
| `mockup-regras-mecanica-v2.html` | v2 — cards por componente + esquema do módulo (descartada) |
| `mockup-regras-mecanica-v3.html` | v3 — frase editável + matriz de impacto (descartada) |
| `mockup-regras-blocos.html` | v4 — construtor de blocos (base da solução) |
| `mockup-regras-blocos-v2.html` | v5 — blocos ocultos + montagens condicionais |
| **`mockup-regras-blocos-v3.html`** | **versão final — cor única nos chips, simulação enxuta** |
