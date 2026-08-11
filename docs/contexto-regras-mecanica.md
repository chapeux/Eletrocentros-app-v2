# Contexto — Implementação de Regras "MECÂNICA" (dimensões por módulo)

## Situação atual do projeto

App desktop (pywebview + `main.py`) com frontend em `frontend/index.html` / `app.js` / `style.css`,
configurado via dois arquivos externos:

- `config.json`: listas de seletores (tipo de estrutura, planpin, tipomaq, incêndio, segurança,
  complexidade, planejadores) + regras de aplicação (`estruturasSemModulo`, `estruturasSemValorMec`,
  `estruturasSemValorEletr`).
- `regras.json`: array de `área` → `campos` → sub-abas `H` (horas) e `DUR` (duração), cada uma com
  uma `base` (forma de cálculo) e uma lista de `condicoes` (adicionais condicionais).

Formas de cálculo já suportadas no motor (`valorBase()` em `app.js`):
`constante`, `multiplicativa`, `degrau_fixo`, `aditiva`, `tabela` (8 posições, 1 a 8 módulos,
aceita expressão com variável `H` via `evalExpr`), `derivado_h` (pipeline de `etapas`: dividir,
multiplicar, somar, subtrair, arredondar).

O motor recebe hoje: `valorBase(base, mod, campoObj, flagsAtivos, vH)` — onde `mod` é **apenas a
quantidade de módulos** (inteiro). As dimensões de cada módulo (`comp` × `larg`) já são capturadas
na tela (inputs `.mod-comp` / `.mod-larg` dentro de `moduleInputs`), mas **não chegam** ao motor de
cálculo — só são usadas hoje para validação visual (LED obrigatório/preenchido).

## Áreas já existentes em `regras.json`

- `ENGENHARIA MECÂNICA`
- `ENGENHARIA ELÉTRICA`

Ambas calculam H/DUR só a partir da **quantidade de módulos**, sem depender das dimensões.

## Próxima etapa: novas áreas a criar

1. **MECÂNICA** (regras "Mecânica 1 a 8")
   - **Depende da quantidade de módulos inserida pelo usuário** — as regras variam conforme o
     número de módulos (1 a 8), igual ao padrão de `tabela` já usado.
   - **Também precisa das dimensões (comp × larg) de cada módulo**, que hoje não entram no cálculo.
   - Essa é a única área que exige mudança no motor.

2. **ACESSÓRIOS**
   - Não depende de dimensões — funciona só por quantidade de módulos, igual à Engenharia atual.
   - Implementação trivial: nova área no `regras.json`, reaproveitando as formas existentes
     (`tabela`, `multiplicativa`, `constante`, `derivado_h`). Sem mudança de código.

3. **ELETROMECÂNICA**
   - Mesma situação de Acessórios: só quantidade de módulos, sem dimensões, sem mudança de código.

## Plano de implementação para MECÂNICA

### a) Capturar as dimensões agregadas dos módulos ativos

```js
function getModuleDims() {
  var dims = [];
  moduleInputs.forEach(function (row, i0) {
    if (i0 >= parseInt(selVal('nrmodulos') || '0', 10)) return;
    var c = parseFloat((row.querySelector('.mod-comp').value || '0').replace(',', '.')) || 0;
    var l = parseFloat((row.querySelector('.mod-larg').value || '0').replace(',', '.')) || 0;
    dims.push({ comp: c, larg: l });
  });
  var C = dims.reduce(function (s, d) { return s + d.comp; }, 0);          // comprimento total
  var A = dims.reduce(function (s, d) { return s + d.comp * d.larg; }, 0); // área total (Σ comp×larg)
  var L = dims.length ? dims.reduce(function (s, d) { return s + d.larg; }, 0) / dims.length : 0; // largura média
  return { dims: dims, C: C, L: L, A: A };
}
```

> Agregação sugerida: `A` (área total = soma de comp×larg de cada módulo) é o candidato mais
> provável para ditar tempo de mecânica, mas `C` (comprimento total) e `L` (largura média) também
> ficam disponíveis. Confirmar com o usuário do domínio qual variável cada campo `MEC*` deve usar.

### b) Estender `evalExpr` para aceitar C, L, A além de H

Hoje: `evalExpr(base.valores[mod-1], hVal)` só injeta a variável `H`.

```js
function evalExpr(val, vH, dimsObj) {
  // ...
  str = str
    .replace(/\bH\b/g, hVal)
    .replace(/\bC\b/g, dimsObj.C)
    .replace(/\bL\b/g, dimsObj.L)
    .replace(/\bA\b/g, dimsObj.A);
  // resto do parser/eval permanece igual
}
```

### c) Marcar a área como dependente de dimensão no `regras.json`

Flag no nível da **área** (não repetir em cada campo):

```json
{
  "area": "MECÂNICA",
  "usa_dimensoes": true,
  "campos": {
    "MEC1": {
      "H": {
        "base": {
          "forma": "tabela",
          "valores": ["A*0.9", "A*0.85", "A*0.8", "A*0.78", "A*0.75", "A*0.73", "A*0.7", "A*0.68"]
        },
        "condicoes": []
      },
      "DUR": {
        "base": { "forma": "derivado_h", "etapas": [{ "tipo": "dividir", "valor": 7.92 }, { "tipo": "arredondar", "modo": "cima" }] },
        "condicoes": []
      }
    }
  }
}
```

Continua sendo a mesma `forma: "tabela"` de 8 posições (1 a 8 módulos) já usada em outras áreas —
só que agora cada posição pode referenciar `A`, `C`, `L`, além de `H`. Reaproveita o editor visual
e o `evalExpr` existentes, sem criar uma "forma" nova do zero.

### d) Em `calcValor` / no ponto que chama `valorBase`

Quando `areaObj.usa_dimensoes === true`, calcular `getModuleDims()` e repassar para `valorBase`
(que repassa para `evalExpr`). Quando `false` (Engenharia, Acessórios, Eletromecânica), não
calcular/ignorar — zero impacto nessas áreas.

### e) UI do editor de regras

Quando a área selecionada tiver `usa_dimensoes: true`, mostrar uma dica abaixo do grid "Tabela por
Módulo": *"Variáveis disponíveis: H (horas), C (comprimento total), L (largura média), A (área
total)"*. Só texto de apoio, sem mudar o layout existente.

## Estruturas sem módulo (lembrete de `config.json`)

`estruturasSemModulo`: Container Solar, Skid (mecânica), Skid (com elétrica), Pilotis,
ESSW (mecânica), ESSW (elétrica), Serviço Engenharia — essas não expandem o bloco de módulos e,
portanto, nunca vão acionar `usa_dimensoes` (não há `comp`/`larg` para elas).

## Tarefas pendentes

- [ ] Confirmar, campo a campo de MECÂNICA (MEC1...MEC8), qual variável (A, C ou L, ou combinação)
      cada um deve usar.
- [ ] Implementar `getModuleDims()`, estender `evalExpr` e o ponto de chamada em `calcValor`.
- [ ] Adicionar `usa_dimensoes: true` na área `MECÂNICA` em `regras.json`.
- [ ] Criar áreas `ACESSÓRIOS` e `ELETROMECÂNICA` em `regras.json` (sem mudança de código).
- [ ] Adicionar dica de variáveis disponíveis no editor de regras quando `usa_dimensoes` for `true`.
- [ ] Popular os campos MEC1–MEC8, ACC(s) e EMC(s) com os valores reais de horas/duração.
