# Código VBA completo — Planilha de Tempos Eletrocentros V22.1

Extraído do `vbaProject.bin` (parser CFB + descompressão MS-OVBA feitos manualmente, sem alterações de conteúdo). Organizado por módulo e, dentro de cada módulo, por Sub/Function, para facilitar navegação e uso como contexto em IDEs.

## Índice de módulos

- [`EstaPasta_de_trabalho`](#estapasta-de-trabalho) — 1 procedimento(s)
- [`template`](#template) — 0 procedimento(s)
- [`tempos`](#tempos) — 0 procedimento(s)
- [`resultado`](#resultado) — 0 procedimento(s)
- [`seletor_template`](#seletor-template) — 0 procedimento(s)
- [`Planilha1`](#planilha1) — 0 procedimento(s)
- [`Planilha2`](#planilha2) — 0 procedimento(s)
- [`a_constants`](#a-constants) — 1 procedimento(s)
- [`f_functions`](#f-functions) — 3 procedimento(s)
- [`x_clear`](#x-clear) — 1 procedimento(s)
- [`FormularioDados`](#formulariodados) — 23 procedimento(s)
- [`Horas_Orçadas_Salas`](#horas-orçadas-salas) — 0 procedimento(s)
- [`Alterar_CTs_Solar`](#alterar-cts-solar) — 1 procedimento(s)
- [`Plan_Tempos`](#plan-tempos) — 17 procedimento(s)

---


## EstaPasta_de_trabalho

### Declarações do módulo (topo do arquivo)

```vb
Attribute VB_Name = "EstaPasta_de_trabalho"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
```

### Procedimentos

#### `Workbook_Open` (Sub)

```vb
Private Sub Workbook_Open()

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
End Sub
```


## template

### Declarações do módulo (topo do arquivo)

```vb
Attribute VB_Name = "template"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
```


## tempos

### Declarações do módulo (topo do arquivo)

```vb
Attribute VB_Name = "tempos"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
```


## resultado

### Declarações do módulo (topo do arquivo)

```vb
Attribute VB_Name = "resultado"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
```


## seletor_template

### Declarações do módulo (topo do arquivo)

```vb
Attribute VB_Name = "seletor_template"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
```


## Planilha1

### Declarações do módulo (topo do arquivo)

```vb
Attribute VB_Name = "Planilha1"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
```


## Planilha2

### Declarações do módulo (topo do arquivo)

```vb
Attribute VB_Name = "Planilha2"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
```


## a_constants

### Declarações do módulo (topo do arquivo)

```vb
Attribute VB_Name = "a_constants"
Public wb As Workbook
Public ws_fields As Worksheet
Public changing_tab As Boolean
```

### Procedimentos

#### `set_wb` (Sub)

```vb
Sub set_wb()

    If wb Is Nothing Then
        Set wb = ThisWorkbook
        Set ws_fields = wb.Sheets("FieldValues")
    End If

End Sub
```


## f_functions

### Declarações do módulo (topo do arquivo)

```vb
Attribute VB_Name = "f_functions"
```

### Procedimentos

#### `get_controlStatus` (Sub)

```vb
Sub get_controlStatus(form As UserForm, sheetName)

    Dim ctrl As Control
    Dim i As Long
    Dim ws As Worksheet

    Call clearFieldValues(sheetName)
    Set ws = wb.Sheets(sheetName)
    
    i = 2
    For Each ctrl In form.Controls
        
        On Error Resume Next
        ws.Cells(i, 1).Value = ctrl.Name
        ws.Cells(i, 2).Value = TypeName(ctrl)
        ws.Cells(i, 3).Value = ctrl.Caption
        ws.Cells(i, 4).Value = ctrl.Visible
        ws.Cells(i, 5).Value = ctrl.Enabled
        ws.Cells(i, 6).Value = ctrl.Value
        i = i + 1
    Next ctrl

End Sub
```

#### `set_controlStatus` (Sub)

```vb
Sub set_controlStatus(form As UserForm, sheetName)

    Dim ctrl As Control
    Dim i As Long
    Dim rng As Range
    Dim rng_aux As Range
    Dim lastrow As Long
    Dim ctrlName As String
    Dim ws As Worksheet
    
    Call set_wb
    
    Set ws = wb.Sheets(sheetName)

    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    If lastrow <= 1 Then: Exit Sub
    
    Set rng = ws.Range("A2:A" & lastrow)

    For Each ctrl In form.Controls
        ctrlName = ctrl.Name
        Set rng_aux = rng.Find(ctrlName, LookAt:=xlWhole, MatchCase:=True)
        If Not rng_aux Is Nothing Then
            On Error Resume Next
            Application.EnableEvents = False
            ctrl.Caption = ws.Cells(rng_aux.Row, 3).Value
            ctrl.Visible = ws.Cells(rng_aux.Row, 4).Value
            ctrl.Enabled = ws.Cells(rng_aux.Row, 5).Value
            ctrl.Value = ws.Cells(rng_aux.Row, 6).Value
            Application.EnableEvents = True
        End If
    Next ctrl

End Sub
```

#### `turn_all_on` (Sub)

```vb
Sub turn_all_on()

    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub
```


## x_clear

### Declarações do módulo (topo do arquivo)

```vb
Attribute VB_Name = "x_clear"
```

### Procedimentos

#### `clearFieldValues` (Sub)

```vb
Sub clearFieldValues(sheetName)
    Dim ws As Worksheet
    Call set_wb
    
    Set ws = wb.Sheets(sheetName)
    
    ws.Range("A2:Z999999").Clear

End Sub
```


## FormularioDados

### Declarações do módulo (topo do arquivo)

```vb
Attribute VB_Name = "FormularioDados"
Attribute VB_Base = "0{D984771C-9AEC-44FB-9AF6-D76E1013C569}{7E09DEAA-B7C5-49DA-8AB7-173855A913B9}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
```

### Procedimentos

#### `Betim1310_Click` (Sub)

```vb
Private Sub Betim1310_Click()

If Application.EnableEvents = False Then: Exit Sub

If FormularioDados.Betim1310.Value = True Then
    FormularioDados.Agrupamento.Visible = False
    FormularioDados.Agrupamento.Value = False
    FormularioDados.Label31.Visible = False
    FormularioDados.proBTI.Visible = False
    FormularioDados.Label33.Visible = False
Else
    FormularioDados.Agrupamento.Visible = True
    FormularioDados.Label31.Visible = True
    FormularioDados.proBTI.Visible = True
    FormularioDados.Label33.Visible = True
End If

If FormularioDados.tipoestrutura.Value = "Container Solar" And FormularioDados.Betim1310.Value = True Then
    FormularioDados.tipomaq.Value = "Não possui"
    FormularioDados.incendio.Value = "Somente Infra"
    FormularioDados.seguranca.Value = "Não possui"
    FormularioDados.complexidade.Value = "Média"
End If

End Sub
```

#### `bt_sair_Click` (Sub)

```vb
Private Sub bt_sair_Click()
    FormularioDados.Hide
    Call get_controlStatus(FormularioDados, "FieldValues")
End Sub
```

#### `calcular_tempos_Click` (Sub)

```vb
Private Sub calcular_tempos_Click()

Call inicio_programa
Call get_controlStatus(FormularioDados, "FieldValues")

End Sub
```

#### `incendio_Change` (Sub)

```vb
Private Sub incendio_Change()

If Application.EnableEvents = False Then: Exit Sub

If incendio.Value = "Não aplicável" And FormularioDados.tipoestrutura.Value <> "Pilotis" And FormularioDados.tipoestrutura.Value <> "Skid (mecânica)" And FormularioDados.tipoestrutura.Value <> "Skid (com elétrica)" And FormularioDados.tipoestrutura.Value <> "ESSW (mecânica)" And FormularioDados.tipoestrutura.Value <> "ESSW (elétrica)" And FormularioDados.tipoestrutura.Value <> "Serviço Engenharia" Then
    MsgBox "O tipo de sistema de incêndio precisa ser determinado para o tipo de eletrocentro selecionado. Favor verificar.", vbOKOnly
    incendio.Value = ""
    Exit Sub
End If

If incendio.Value <> "Não aplicável" And (tipoestrutura.Value = "ESSW (elétrica)" Or tipoestrutura.Value = "ESSW (mecânica)" Or tipoestrutura.Value = "Skid (mecânica)" Or tipoestrutura.Value = "Skid (com elétrica)" Or tipoestrutura.Value = "Pilotis" Or FormularioDados.tipoestrutura.Value = "Serviço Engenharia") Then
    MsgBox "Campo não aplicável para o tipo de eletrocentro selecionado.", vbOKOnly
    incendio.Value = "Não aplicável"
    Exit Sub
End If

End Sub
```

#### `complexidade_Change` (Sub)

```vb
Private Sub complexidade_Change()

If Application.EnableEvents = False Then: Exit Sub

If complexidade.Value = "Não aplicável" And FormularioDados.tipoestrutura.Value <> "Pilotis" And FormularioDados.tipoestrutura.Value <> "Skid (mecânica)" And FormularioDados.tipoestrutura.Value <> "Skid (com elétrica)" And FormularioDados.tipoestrutura.Value <> "ESSW (mecânica)" And FormularioDados.tipoestrutura.Value <> "ESSW (elétrica)" And FormularioDados.tipoestrutura.Value <> "Serviço Engenharia" Then
    MsgBox "A complexidade dos equipamentos precisa ser determinada para o tipo de eletrocentro selecionado. Favor verificar.", vbOKOnly
    complexidade.Value = ""
    Exit Sub
End If
If complexidade.Value = "Não aplicável" Then
    nrcolunas.Value = 0
    nrcolunas.Enabled = False
    '#DESABILITADO PAINEIS INTERLIGAÇÃO
    'paineisint.Value = 0
    'paineisint.Enabled = False
Else
    nrcolunas.Value = 0
    nrcolunas.Enabled = True
    '#DESABILITADO PAINEIS INTERLIGAÇÃO
    'paineisint.Value = 0
    'paineisint.Enabled = True
End If

If complexidade.Value <> "Não aplicável" And (tipoestrutura.Value = "ESSW (elétrica)" Or tipoestrutura.Value = "ESSW (mecânica)" Or tipoestrutura.Value = "Skid (mecânica)" Or tipoestrutura.Value = "Skid (com elétrica)" Or tipoestrutura.Value = "Pilotis") Then
    MsgBox "Campo não aplicável para o tipo de eletrocentro selecionado.", vbOKOnly
    complexidade.Value = "Não aplicável"
    Exit Sub
End If

End Sub
```

#### `lb_N_OV_Click` (Sub)

```vb
Private Sub lb_N_OV_Click()

End Sub
```

#### `Mecanica_Click` (Sub)

```vb
Private Sub Mecanica_Click()

End Sub
```

#### `nrmodulos_Change` (Sub)

```vb
Private Sub nrmodulos_Change()

If Application.EnableEvents = False Then: Exit Sub

If nrmodulos.Value = "1 Módulo" And tipoestrutura.Value <> "Container Solar" And tipoestrutura.Value <> "Pilotis" And tipoestrutura.Value <> "Skid (mecânica)" And tipoestrutura.Value <> "Skid (com elétrica)" And tipoestrutura.Value <> "ESSW (mecânica)" And tipoestrutura.Value <> "ESSW (elétrica)" And tipoestrutura.Value <> "Serviço Engenharia" Then
    modulo1.Enabled = True
    modulo2.Value = 0
    modulo2.Enabled = False
    modulo3.Value = 0
    modulo3.Enabled = False
    modulo4.Value = 0
    modulo4.Enabled = False
    modulo5.Value = 0
    modulo5.Enabled = False
    modulo6.Value = 0
    modulo6.Enabled = False
    modulo7.Value = 0
    modulo7.Enabled = False
    modulo8.Value = 0
    modulo8.Enabled = False
    largmodulo1.Enabled = True
    largmodulo2.Value = 0
    largmodulo2.Enabled = False
    largmodulo3.Value = 0
    largmodulo3.Enabled = False
    largmodulo4.Value = 0
    largmodulo4.Enabled = False
    largmodulo5.Value = 0
    largmodulo5.Enabled = False
    largmodulo6.Value = 0
    largmodulo6.Enabled = False
    largmodulo7.Value = 0
    largmodulo7.Enabled = False
    largmodulo8.Value = 0
    largmodulo8.Enabled = False
    chkPeDireito.Enabled = True
End If

If nrmodulos.Value = "2 Módulos" And tipoestrutura.Value <> "Container Solar" And tipoestrutura.Value <> "Pilotis" And tipoestrutura.Value <> "Skid (mecânica)" And tipoestrutura.Value <> "Skid (com elétrica)" And tipoestrutura.Value <> "ESSW (mecânica)" And tipoestrutura.Value <> "ESSW (elétrica)" And tipoestrutura.Value <> "Serviço Engenharia" Then
    modulo1.Enabled = True
    modulo2.Enabled = True
    modulo3.Value = 0
    modulo3.Enabled = False
    modulo4.Value = 0
    modulo4.Enabled = False
    modulo5.Value = 0
    modulo5.Enabled = False
    modulo6.Value = 0
    modulo6.Enabled = False
    modulo7.Value = 0
    modulo7.Enabled = False
    modulo8.Value = 0
    modulo8.Enabled = False
    largmodulo1.Enabled = True
    largmodulo2.Enabled = True
    largmodulo3.Value = 0
    largmodulo3.Enabled = False
    largmodulo4.Value = 0
    largmodulo4.Enabled = False
    largmodulo5.Value = 0
    largmodulo5.Enabled = False
    largmodulo6.Value = 0
    largmodulo6.Enabled = False
    largmodulo7.Value = 0
    largmodulo7.Enabled = False
    largmodulo8.Value = 0
    largmodulo8.Enabled = False
    chkPeDireito.Enabled = True
End If

If nrmodulos.Value = "3 Módulos" And tipoestrutura.Value <> "Container Solar" And tipoestrutura.Value <> "Pilotis" And tipoestrutura.Value <> "Skid (mecânica)" And tipoestrutura.Value <> "Skid (com elétrica)" And tipoestrutura.Value <> "ESSW (mecânica)" And tipoestrutura.Value <> "ESSW (elétrica)" And tipoestrutura.Value <> "Serviço Engenharia" Then
    modulo1.Enabled = True
    modulo2.Enabled = True
    modulo3.Enabled = True
    modulo4.Value = 0
    modulo4.Enabled = False
    modulo5.Value = 0
    modulo5.Enabled = False
    modulo6.Value = 0
    modulo6.Enabled = False
    modulo7.Value = 0
    modulo7.Enabled = False
    modulo8.Value = 0
    modulo8.Enabled = False
    largmodulo1.Enabled = True
    largmodulo2.Enabled = True
    largmodulo3.Enabled = True
    largmodulo4.Value = 0
    largmodulo4.Enabled = False
    largmodulo5.Value = 0
    largmodulo5.Enabled = False
    largmodulo6.Value = 0
    largmodulo6.Enabled = False
    largmodulo7.Value = 0
    largmodulo7.Enabled = False
    largmodulo8.Value = 0
    largmodulo8.Enabled = False
    chkPeDireito.Enabled = True
End If

If nrmodulos.Value = "4 Módulos" And tipoestrutura.Value <> "Container Solar" And tipoestrutura.Value <> "Pilotis" And tipoestrutura.Value <> "Skid (mecânica)" And tipoestrutura.Value <> "Skid (com elétrica)" And tipoestrutura.Value <> "ESSW (mecânica)" And tipoestrutura.Value <> "ESSW (elétrica)" And tipoestrutura.Value <> "Serviço Engenharia" Then
    modulo1.Enabled = True
    modulo2.Enabled = True
    modulo3.Enabled = True
    modulo4.Enabled = True
    modulo5.Value = 0
    modulo5.Enabled = False
    modulo6.Value = 0
    modulo6.Enabled = False
    modulo7.Value = 0
    modulo7.Enabled = False
    modulo8.Value = 0
    modulo8.Enabled = False
    largmodulo1.Enabled = True
    largmodulo2.Enabled = True
    largmodulo3.Enabled = True
    largmodulo4.Enabled = True
    largmodulo5.Value = 0
    largmodulo5.Enabled = False
    largmodulo6.Value = 0
    largmodulo6.Enabled = False
    largmodulo7.Value = 0
    largmodulo7.Enabled = False
    largmodulo8.Value = 0
    largmodulo8.Enabled = False
    chkPeDireito.Enabled = True
End If

If nrmodulos.Value = "5 Módulos" And tipoestrutura.Value <> "Container Solar" And tipoestrutura.Value <> "Pilotis" And tipoestrutura.Value <> "Skid (mecânica)" And tipoestrutura.Value <> "Skid (com elétrica)" And tipoestrutura.Value <> "ESSW (mecânica)" And tipoestrutura.Value <> "ESSW (elétrica)" And tipoestrutura.Value <> "Serviço Engenharia" Then
    modulo1.Enabled = True
    modulo2.Enabled = True
    modulo3.Enabled = True
    modulo4.Enabled = True
    modulo5.Enabled = True
    modulo6.Value = 0
    modulo6.Enabled = False
    modulo7.Value = 0
    modulo7.Enabled = False
    modulo8.Value = 0
    modulo8.Enabled = False
    largmodulo1.Enabled = True
    largmodulo2.Enabled = True
    largmodulo3.Enabled = True
    largmodulo4.Enabled = True
    largmodulo5.Enabled = True
    largmodulo6.Value = 0
    largmodulo6.Enabled = False
    largmodulo7.Value = 0
    largmodulo7.Enabled = False
    largmodulo8.Value = 0
    largmodulo8.Enabled = False
    chkPeDireito.Enabled = True
End If

If nrmodulos.Value = "6 Módulos" And tipoestrutura.Value <> "Container Solar" And tipoestrutura.Value <> "Pilotis" And tipoestrutura.Value <> "Skid (mecânica)" And tipoestrutura.Value <> "Skid (com elétrica)" And tipoestrutura.Value <> "ESSW (mecânica)" And tipoestrutura.Value <> "ESSW (elétrica)" And tipoestrutura.Value <> "Serviço Engenharia" Then
    modulo1.Enabled = True
    modulo2.Enabled = True
    modulo3.Enabled = True
    modulo4.Enabled = True
    modulo5.Enabled = True
    modulo6.Enabled = True
    modulo7.Value = 0
    modulo7.Enabled = False
    modulo8.Value = 0
    modulo8.Enabled = False
    largmodulo1.Enabled = True
    largmodulo2.Enabled = True
    largmodulo3.Enabled = True
    largmodulo4.Enabled = True
    largmodulo5.Enabled = True
    largmodulo6.Enabled = True
    largmodulo7.Value = 0
    largmodulo7.Enabled = False
    largmodulo8.Value = 0
    largmodulo8.Enabled = False
    chkPeDireito.Enabled = True
End If

If nrmodulos.Value = "7 Módulos" And tipoestrutura.Value <> "Container Solar" And tipoestrutura.Value <> "Pilotis" And tipoestrutura.Value <> "Skid (mecânica)" And tipoestrutura.Value <> "Skid (com elétrica)" And tipoestrutura.Value <> "ESSW (mecânica)" And tipoestrutura.Value <> "ESSW (elétrica)" And tipoestrutura.Value <> "Serviço Engenharia" Then
    modulo1.Enabled = True
    modulo2.Enabled = True
    modulo3.Enabled = True
    modulo4.Enabled = True
    modulo5.Enabled = True
    modulo6.Enabled = True
    modulo7.Enabled = True
    modulo8.Value = 0
    modulo8.Enabled = False
    largmodulo1.Enabled = True
    largmodulo2.Enabled = True
    largmodulo3.Enabled = True
    largmodulo4.Enabled = True
    largmodulo5.Enabled = True
    largmodulo6.Enabled = True
    largmodulo7.Enabled = True
    largmodulo8.Value = 0
    largmodulo8.Enabled = False
    chkPeDireito.Enabled = True
End If

If nrmodulos.Value = "8 Módulos" And tipoestrutura.Value <> "Container Solar" And tipoestrutura.Value <> "Pilotis" And tipoestrutura.Value <> "Skid (mecânica)" And tipoestrutura.Value <> "Skid (com elétrica)" And tipoestrutura.Value <> "ESSW (mecânica)" And tipoestrutura.Value <> "ESSW (elétrica)" And tipoestrutura.Value <> "Serviço Engenharia" Then
    modulo1.Enabled = True
    modulo2.Enabled = True
    modulo3.Enabled = True
    modulo4.Enabled = True
    modulo5.Enabled = True
    modulo6.Enabled = True
    modulo7.Enabled = True
    modulo8.Enabled = True
    largmodulo1.Enabled = True
    largmodulo2.Enabled = True
    largmodulo3.Enabled = True
    largmodulo4.Enabled = True
    largmodulo5.Enabled = True
    largmodulo6.Enabled = True
    largmodulo7.Enabled = True
    largmodulo8.Enabled = True
    chkPeDireito.Enabled = True
End If

End Sub
```

#### `planpin_Change` (Sub)

```vb
Private Sub planpin_Change()

If Application.EnableEvents = False Then: Exit Sub

If planpin.Value = "WAU-ELETRO-04" And tipoestrutura.Value <> "Container Solar" Then
    MsgBox "O plano de pintura WAU-ELETRO-04 é aplicável apenas para Container Solar.", vbOKOnly
    planpin.Value = ""
End If

If planpin.Value = "Não aplicável" And tipoestrutura.Value <> "ESSW (elétrica)" And tipoestrutura.Value <> "Serviço Engenharia" Then
    MsgBox "O plano de pintura deve ser informado para o tipo de eletrocentro informado. Favor verificar.", vbOKOnly
    planpin.Value = ""
    Exit Sub
End If

If planpin.Value <> "Não aplicável" And (tipoestrutura.Value = "ESSW (elétrica)" Or tipoestrutura.Value = "Serviço Engenharia") Then
    MsgBox "Campo não aplicável para o tipo de eletrocentro selecionado.", vbOKOnly
    planpin.Value = "Não aplicável"
    Exit Sub
End If

End Sub
```

#### `proBTI_Click` (Sub)

```vb
Private Sub proBTI_Click()

If Application.EnableEvents = False Then: Exit Sub

If FormularioDados.proBTI.Value = True Then
    FormularioDados.fr_outrosCentros.Visible = False
    If FormularioDados.chkMaterial.Value = True Then
        FormularioDados.Label32.Visible = True
        FormularioDados.txtMateriaavo.Visible = True
    End If
Else
    FormularioDados.fr_outrosCentros.Visible = True
    FormularioDados.Label32.Visible = False
    FormularioDados.txtMateriaavo.Visible = False
End If
 
End Sub
```

#### `seguranca_Change` (Sub)

```vb
Private Sub seguranca_Change()

If Application.EnableEvents = False Then: Exit Sub

If seguranca.Value = "Não aplicável" And tipoestrutura.Value <> "ESSW (elétrica)" And tipoestrutura.Value <> "ESSW (mecânica)" And tipoestrutura.Value <> "Skid (mecânica)" And tipoestrutura.Value <> "Skid (com elétrica)" And tipoestrutura.Value <> "Pilotis" And tipoestrutura.Value <> "Serviço Engenharia" Then
    MsgBox "O tipo de sistema de segurança deve ser informado para o tipo de eletrocentro informado. Caso não possua, selecionar 'Não possui'. Favor verificar.", vbOKOnly
    seguranca.Value = ""
    Exit Sub
End If

If seguranca.Value <> "Não aplicável" And (tipoestrutura.Value = "ESSW (elétrica)" Or tipoestrutura.Value = "ESSW (mecânica)" Or tipoestrutura.Value = "Skid (mecânica)" Or tipoestrutura.Value = "Skid (com elétrica)" Or tipoestrutura.Value = "Pilotis" Or tipoestrutura.Value = "Serviço Engenharia") Then
    MsgBox "Campo não aplicável para o tipo de eletrocentro selecionado.", vbOKOnly
    seguranca.Value = "Não aplicável"
    Exit Sub
End If

End Sub
```

#### `SemEng_Click` (Sub)

```vb
Private Sub SemEng_Click()

'If tipoestrutura.Value <> "Container Solar" And SemEng.Value = True Then
'    MsgBox "Opção válida apenas para Container Solar.", vbOKOnly
'    SemEng.Value = False
'    Exit Sub
'End If

End Sub
```

#### `TextBox1_Change` (Sub)

```vb
Private Sub TextBox1_Change()

End Sub
```

#### `TabStrip1_Change` (Sub)

```vb
Private Sub TabStrip1_Change()
    
    If Application.EnableEvents = False Then: Exit Sub
    
    Dim tab_atual As Integer
    tab_atual = Me.TabStrip1.Value
    
    'RESETA OS VALORES PARA O ESTADO INICIAL
    Call set_controlStatus(FormularioDados, "InitialValues")
    Application.EnableEvents = False
    Me.TabStrip1.Value = tab_atual
    
    
    Select Case tab_atual
    
        Case 0 'PCP
            Me.fr_Container_Marítimo.Visible = True
            Me.fr_SAP.Visible = True
            Me.fr_sapInfo.Visible = True
            Me.fr_horasOrcadas.Visible = True
            Me.fr_outrosCentros.Visible = True

        Case 1 'VENDAS
            Me.fr_Container_Marítimo.Visible = False
            Me.fr_SAP.Visible = False
            Me.fr_sapInfo.Visible = False
            Me.fr_horasOrcadas.Visible = False
            Me.fr_outrosCentros.Visible = False
    End Select
    
    Application.EnableEvents = True
    
End Sub
```

#### `tipoestrutura_Change` (Sub)

```vb
Private Sub tipoestrutura_Change()

If Application.EnableEvents = False Then: Exit Sub

If tipoestrutura.Value = "Container Solar" Or tipoestrutura.Value = "Pilotis" Or tipoestrutura.Value = "Skid (mecânica)" Or tipoestrutura.Value = "Skid (com elétrica)" Or tipoestrutura.Value = "ESSW (mecânica)" Or tipoestrutura.Value = "ESSW (elétrica)" Or tipoestrutura.Value = "Serviço Engenharia" Then
    modulo1.Value = 0
    modulo1.Enabled = False
    modulo2.Value = 0
    modulo2.Enabled = False
    modulo3.Value = 0
    modulo3.Enabled = False
    modulo4.Value = 0
    modulo4.Enabled = False
    modulo5.Value = 0
    modulo5.Enabled = False
    modulo6.Value = 0
    modulo6.Enabled = False
    modulo7.Value = 0
    modulo7.Enabled = False
    modulo8.Value = 0
    modulo8.Enabled = False
    largmodulo1.Value = 0
    largmodulo1.Enabled = False
    largmodulo2.Value = 0
    largmodulo2.Enabled = False
    largmodulo3.Value = 0
    largmodulo3.Enabled = False
    largmodulo4.Value = 0
    largmodulo4.Enabled = False
    largmodulo5.Value = 0
    largmodulo5.Enabled = False
    largmodulo6.Value = 0
    largmodulo6.Enabled = False
    largmodulo7.Value = 0
    largmodulo7.Enabled = False
    largmodulo8.Value = 0
    largmodulo8.Enabled = False
    If tipoestrutura.Value <> "Serviço Engenharia" Then nrmodulos.Value = "1 Módulo"
    If tipoestrutura.Value <> "Serviço Engenharia" Then nrmodulos.Enabled = False
    If tipoestrutura.Value = "Container Solar" Then planpin.Value = "WAU-ELETRO-04"
    If tipoestrutura.Value <> "Container Solar" Then incendio.Value = "Não aplicável"
    If tipoestrutura.Value <> "Container Solar" Then complexidade.Value = "Não aplicável"
    If tipoestrutura.Value = "ESSW (elétrica)" Or tipoestrutura.Value = "Serviço Engenharia" Then planpin.Value = "Não aplicável"
    If tipoestrutura.Value <> "Container Solar" Then tipomaq.Value = "Não aplicável"
    If tipoestrutura.Value <> "Container Solar" Then seguranca.Value = "Não aplicável"
    If tipoestrutura.Value = "Container Solar" And incendio.Value = "Não aplicável" Then incendio.Value = ""
    If tipoestrutura.Value = "Container Solar" And complexidade.Value = "Não aplicável" Then complexidade.Value = ""
    If tipoestrutura.Value = "Container Solar" And tipomaq.Value = "Não aplicável" Then tipomaq.Value = ""
    If tipoestrutura.Value = "Container Solar" And seguranca.Value = "Não aplicável" Then seguranca.Value = ""
    
    If tipoestrutura.Value = "Serviço Engenharia" Then
        FormularioDados.txtMaterial.Visible = False
        FormularioDados.Label16.Visible = False
    End If
    
    If tipoestrutura.Value <> "Container Solar" Then
        testesw.Enabled = False
        testesw.Value = False
    End If
    If tipoestrutura.Value = "Container Solar" Then
        testesw.Enabled = True
    End If
    whitemartins.Value = False
    whitemartins.Enabled = False
    If tipoestrutura.Value <> "Serviço Engenharia" Then chkPeDireito.Value = False
    If tipoestrutura.Value <> "Serviço Engenharia" Then chkPeDireito.Enabled = False
    If tipoestrutura.Value = "Skid (mecânica)" Or tipoestrutura.Value = "Skid (com elétrica)" Or tipoestrutura.Value = "Serviço Engenharia" Then
        trafooleo.Value = False
        trafooleo.Enabled = True
    Else
        trafooleo.Value = False
        trafooleo.Enabled = False
    End If
    
    If FormularioDados.tipoestrutura.Value = "Skid (com elétrica)" Or FormularioDados.tipoestrutura.Value = "Container Solar" Then
        nrcolunas.Value = 0
        nrcolunas.Enabled = True
    Else
        nrcolunas.Value = 0
        nrcolunas.Enabled = False
    End If
 
    
    If tipoestrutura.Value <> "Container Solar" And planpin.Value = "WAU-ELETRO-04" Then planpin.Value = ""
    If tipoestrutura.Value <> "ESSW (elétrica)" And tipoestrutura.Value <> "Serviço Engenharia" And planpin.Value = "Não aplicável" Then planpin.Value = ""
    If (tipoestrutura.Value = "ESSW (mecânica)" Or tipoestrutura.Value = "Skid (mecânica)" Or tipoestrutura.Value = "Pilotis") And FormularioDados.chkCamposUsuario.Value = True Then
        FormularioDados.txtValorEletr.Value = ""
        FormularioDados.txtValorEletr.Visible = False
        FormularioDados.Label24.Visible = False
        FormularioDados.txtValor.Visible = True
        FormularioDados.Label14.Visible = True
    End If
    If tipoestrutura.Value = "ESSW (elétrica)" Or tipoestrutura.Value = "Serviço Engenharia" And FormularioDados.chkCamposUsuario.Value = True Then
        FormularioDados.txtValor.Value = ""
        FormularioDados.txtValor.Visible = False
        FormularioDados.Label14.Visible = False
        FormularioDados.txtValorEletr.Visible = True
        FormularioDados.Label24.Visible = True
    End If
    If tipoestrutura.Value = "Skid (com elétrica)" And FormularioDados.chkCamposUsuario.Value = True Then
        FormularioDados.txtValor.Visible = True
        FormularioDados.Label14.Visible = True
        FormularioDados.txtValorEletr.Visible = True
        FormularioDados.Label24.Visible = True
    End If
    If tipoestrutura.Value = "Container Solar" And FormularioDados.chkCamposUsuario.Value = True Then
        FormularioDados.txtValorEletr.Visible = True
        FormularioDados.Label24.Visible = True
        FormularioDados.txtValor.Visible = True
        FormularioDados.Label14.Visible = True
    End If
Else
    If FormularioDados.chkCamposUsuario.Value = True Then
        FormularioDados.txtValorEletr.Visible = True
        FormularioDados.Label24.Visible = True
        FormularioDados.txtValor.Visible = True
        FormularioDados.Label14.Visible = True
    End If
    nrmodulos.Enabled = True
    modulo1.Enabled = True
    largmodulo1.Enabled = True
    If planpin.Value = "WAU-ELETRO-04" Then
        planpin.Value = ""
    End If
    If incendio.Value = "Não aplicável" Then incendio.Value = ""
    If complexidade.Value = "Não aplicável" Then complexidade.Value = ""
    If planpin.Value = "Não aplicável" Then planpin.Value = ""
    If seguranca.Value = "Não aplicável" Then seguranca.Value = ""
    If tipomaq.Value = "Não aplicável" Then tipomaq.Value = ""
    testesw.Enabled = True
    whitemartins.Enabled = True
    trafooleo.Enabled = True
    chkPeDireito.Enabled = True
End If

If FormularioDados.tipoestrutura.Value = "Container Solar" Then
    FormularioDados.Agrupamento.Visible = True
    FormularioDados.Label31.Visible = True
    FormularioDados.programacaoreles.Visible = True
    FormularioDados.programacaoreles_label.Visible = True
Else
    FormularioDados.Agrupamento.Visible = False
    FormularioDados.Agrupamento.Value = False
    FormularioDados.Label31.Visible = False
    FormularioDados.programacaoreles.Visible = False
    FormularioDados.programacaoreles_label.Visible = False
End If

If FormularioDados.tipoestrutura.Value <> "ESSW (elétrica)" And FormularioDados.tipoestrutura.Value <> "ESSW (mecânica)" And tipoestrutura.Value <> "Serviço Engenharia" Then
    FormularioDados.fr_outrosCentros.Visible = True
Else
    If FormularioDados.Betim1310.Value = True Then
        FormularioDados.tipoestrutura.Value = ""
        MsgBox "Tipo de eletrocentro selecionado não permitido para fabricação em Betim. Favor verificar.", vbOKOnly
        Exit Sub
    Else
        FormularioDados.fr_outrosCentros.Visible = False
    End If
End If

If (FormularioDados.tipoestrutura.Value <> "Container Solar" And FormularioDados.tipoestrutura.Value <> "Skid (com elétrica)" And FormularioDados.Betim1310.Value = True) Or tipoestrutura.Value = "Serviço Engenharia" Then
    FormularioDados.SemEng.Visible = False
    FormularioDados.SemEng.Value = False
    FormularioDados.Label28.Visible = False
Else
    FormularioDados.SemEng.Visible = True
    FormularioDados.Label28.Visible = True
    
    If FormularioDados.Betim1310.Value = False Then
        FormularioDados.proBTI.Visible = True
        FormularioDados.Label33.Visible = True
    Else
        FormularioDados.proBTI.Visible = False
        FormularioDados.Label33.Visible = False
    End If
End If

End Sub
```

#### `tipomaq_Change` (Sub)

```vb
Private Sub tipomaq_Change()

If Application.EnableEvents = False Then: Exit Sub

If tipomaq.Value = "Não aplicável" And tipoestrutura.Value <> "ESSW (elétrica)" And tipoestrutura.Value <> "ESSW (mecânica)" And tipoestrutura.Value <> "Skid (mecânica)" And tipoestrutura.Value <> "Skid (com elétrica)" And tipoestrutura.Value <> "Pilotis" And tipoestrutura.Value <> "Serviço Engenharia" Then
    MsgBox "O tipo de máquina de ar condicionado deve ser informado para o tipo de eletrocentro informado. Caso não possua, selecionar 'Não possui'. Favor verificar.", vbOKOnly
    tipomaq.Value = ""
    Exit Sub
End If

If tipomaq.Value <> "Não aplicável" And (tipoestrutura.Value = "ESSW (elétrica)" Or tipoestrutura.Value = "ESSW (mecânica)" Or tipoestrutura.Value = "Skid (mecânica)" Or tipoestrutura.Value = "Skid (com elétrica)" Or tipoestrutura.Value = "Pilotis" Or tipoestrutura.Value = "Serviço Engenharia") Then
    MsgBox "Campo não aplicável para o tipo de eletrocentro selecionado.", vbOKOnly
    tipomaq.Value = "Não aplicável"
    Exit Sub
End If

If tipomaq.Value = "Não possui" Or tipomaq.Value = "Não aplicável" Then
    qtdmaq.Value = 0
    qtdmaq.Enabled = False
    'dutos.Value = False
    'dutos.Enabled = False
Else
    qtdmaq.Enabled = True
    'dutos.Enabled = True
End If

End Sub
```

#### `btSair_Click` (Sub)

```vb
Private Sub btSair_Click()
    FormularioDados.Hide
End Sub
```

#### `chkCriarDRs_Click` (Sub)

```vb
Private Sub chkCriarDRs_Click()

If Application.EnableEvents = False Then: Exit Sub

If FormularioDados.chkCriarDRs.Value = True Then
    FormularioDados.txtPEP.Visible = True
    FormularioDados.Label23.Visible = True
Else
    If FormularioDados.chkPlanejar.Value = False And FormularioDados.chkCamposUsuario.Value = False And FormularioDados.chkMaterial.Value = False Then
        FormularioDados.txtPEP.Visible = False
        FormularioDados.Label23.Visible = False
        FormularioDados.txtPEP.Value = ""
    End If
End If

End Sub
```

#### `chkPlanejar_Click` (Sub)

```vb
Private Sub chkPlanejar_Click()

If Application.EnableEvents = False Then: Exit Sub

If FormularioDados.chkPlanejar.Value = True Then
    FormularioDados.txtPEP.Visible = True
    FormularioDados.Label23.Visible = True
    FormularioDados.txtDataInicio.Visible = True
    FormularioDados.Label17.Visible = True
Else
    If FormularioDados.chkCriarDRs.Value = False And FormularioDados.chkCamposUsuario.Value = False And FormularioDados.chkMaterial.Value = False Then
        FormularioDados.txtPEP.Visible = False
        FormularioDados.Label23.Visible = False
        FormularioDados.txtPEP.Value = ""
        FormularioDados.txtDataInicio.Visible = False
        FormularioDados.Label17.Visible = False
        FormularioDados.txtDataInicio.Value = ""
    Else
        FormularioDados.txtDataInicio.Visible = False
        FormularioDados.Label17.Visible = False
        FormularioDados.txtDataInicio.Value = ""
    End If
End If

End Sub
```

#### `chkCamposUsuario_Click` (Sub)

```vb
Private Sub chkCamposUsuario_Click()

If Application.EnableEvents = False Then: Exit Sub

If FormularioDados.chkCamposUsuario.Value = True Then
    FormularioDados.txtPEP.Visible = True
    FormularioDados.Label23.Visible = True
    FormularioDados.txtCliente.Visible = True
    FormularioDados.Label12.Visible = True
    If FormularioDados.tipoestrutura.Value <> "ESSW (elétrica)" And tipoestrutura.Value <> "Serviço Engenharia" Then
        FormularioDados.txtValor.Visible = True
        FormularioDados.Label14.Visible = True
    Else
        FormularioDados.txtValor.Visible = False
        FormularioDados.Label14.Visible = False
    End If
    If FormularioDados.tipoestrutura.Value <> "ESSW (mecânica)" Then
        FormularioDados.txtValorEletr.Visible = True
        FormularioDados.Label24.Visible = True
    Else
        FormularioDados.txtValorEletr.Visible = False
        FormularioDados.Label24.Visible = False
    End If
    FormularioDados.txtDataOV.Visible = True
    FormularioDados.Label15.Visible = True
    FormularioDados.cmbPlanejador.Visible = True
    FormularioDados.Label13.Visible = True
    
    FormularioDados.lb_N_itemOV.Visible = True
    FormularioDados.txt_itemOV.Visible = True
    FormularioDados.txt_N_OV.Visible = True
Else
    FormularioDados.txtCliente.Visible = False
    FormularioDados.Label12.Visible = False
    FormularioDados.txtCliente.Value = ""
    FormularioDados.txtValor.Visible = False
    FormularioDados.Label14.Visible = False
    FormularioDados.txtValor.Value = ""
    FormularioDados.txtDataOV.Visible = False
    FormularioDados.Label15.Visible = False
    FormularioDados.txtValorEletr.Value = ""
    FormularioDados.txtValorEletr.Visible = False
    FormularioDados.Label24.Visible = False
    FormularioDados.txtDataOV.Value = ""
    FormularioDados.cmbPlanejador.Visible = False
    FormularioDados.Label13.Visible = False
    FormularioDados.cmbPlanejador.Value = ""
    If FormularioDados.chkPlanejar.Value = False And FormularioDados.chkCriarDRs.Value = False And FormularioDados.chkMaterial.Value = False Then
        FormularioDados.txtPEP.Visible = False
        FormularioDados.Label23.Visible = False
        FormularioDados.txtPEP.Value = ""
    End If
    
    FormularioDados.lb_N_itemOV.Visible = False
    FormularioDados.txt_itemOV.Visible = False
    FormularioDados.txt_N_OV.Visible = False
End If

End Sub
```

#### `chkMaterial_Click` (Sub)

```vb
Private Sub chkMaterial_Click()

If Application.EnableEvents = False Then: Exit Sub

If FormularioDados.chkMaterial.Value = True Then
    If tipoestrutura.Value <> "Serviço Engenharia" Then FormularioDados.txtMaterial.Visible = True
    If tipoestrutura.Value <> "Serviço Engenharia" Then FormularioDados.Label16.Visible = True
    FormularioDados.txtPEP.Visible = True
    FormularioDados.Label23.Visible = True
Else
    FormularioDados.txtMaterial.Visible = False
    FormularioDados.Label16.Visible = False
    FormularioDados.txtMaterial.Value = ""
    If FormularioDados.chkPlanejar.Value = False And FormularioDados.chkCamposUsuario.Value = False And FormularioDados.chkCriarDRs.Value = False Then
        FormularioDados.txtPEP.Visible = False
        FormularioDados.Label23.Visible = False
        FormularioDados.txtPEP.Value = ""
    End If
End If

End Sub
```

#### `UserForm_Initialize` (Sub)

```vb
Private Sub UserForm_Initialize()

'Inicializa menus dos comboboxes

If FormularioDados.tipoestrutura.ListCount <= 0 Then

    FormularioDados.tipoestrutura.Clear
    FormularioDados.tipoestrutura.AddItem "Móvel", 0
    FormularioDados.tipoestrutura.AddItem "Semimóvel", 1
    FormularioDados.tipoestrutura.AddItem "Modular", 2
    FormularioDados.tipoestrutura.AddItem "Fixo", 3
    FormularioDados.tipoestrutura.AddItem "Embarcado", 4
    FormularioDados.tipoestrutura.AddItem "Container Solar", 5
    FormularioDados.tipoestrutura.AddItem "Skid (mecânica)", 6
    FormularioDados.tipoestrutura.AddItem "Skid (com elétrica)", 7
    FormularioDados.tipoestrutura.AddItem "Pilotis", 8
    FormularioDados.tipoestrutura.AddItem "ESSW (mecânica)", 9
    FormularioDados.tipoestrutura.AddItem "ESSW (elétrica)", 10
    FormularioDados.tipoestrutura.AddItem "Serviço Engenharia", 11

    FormularioDados.nrmodulos.Clear
    FormularioDados.nrmodulos.AddItem "1 Módulo", 0
    FormularioDados.nrmodulos.AddItem "2 Módulos", 1
    FormularioDados.nrmodulos.AddItem "3 Módulos", 2
    FormularioDados.nrmodulos.AddItem "4 Módulos", 3
    FormularioDados.nrmodulos.AddItem "5 Módulos", 4
    FormularioDados.nrmodulos.AddItem "6 Módulos", 5
    FormularioDados.nrmodulos.AddItem "7 Módulos", 6
    FormularioDados.nrmodulos.AddItem "8 Módulos", 7
    
    FormularioDados.planpin.Clear
    FormularioDados.planpin.AddItem "WAU-ELETRO-08", 0
    FormularioDados.planpin.AddItem "WAU-ELETRO-09", 1
    FormularioDados.planpin.AddItem "WAU-ELETRO-04", 2
    FormularioDados.planpin.AddItem "Não aplicável", 3

    FormularioDados.tipomaq.Clear
    FormularioDados.tipomaq.AddItem "Split", 0
    FormularioDados.tipomaq.AddItem "Wall Mounted", 1
    FormularioDados.tipomaq.AddItem "Roof Top", 2
    FormularioDados.tipomaq.AddItem "Não possui", 3
    FormularioDados.tipomaq.AddItem "Não aplicável", 4
    
    FormularioDados.incendio.Clear
    FormularioDados.incendio.AddItem "Com combate", 0
    FormularioDados.incendio.AddItem "Com instalações", 1
    FormularioDados.incendio.AddItem "Somente infra", 2
    FormularioDados.incendio.AddItem "Não aplicável", 3

    FormularioDados.seguranca.Clear
    FormularioDados.seguranca.AddItem "CFTV", 0
    FormularioDados.seguranca.AddItem "Controle Acesso", 1
    FormularioDados.seguranca.AddItem "CFTV + Controle Acesso", 2
    FormularioDados.seguranca.AddItem "Não possui", 3
    FormularioDados.seguranca.AddItem "Não aplicável", 4

    FormularioDados.complexidade.Clear
    FormularioDados.complexidade.AddItem "Simples", 0
    FormularioDados.complexidade.AddItem "Médio", 1
    FormularioDados.complexidade.AddItem "Complexo", 2
    FormularioDados.complexidade.AddItem "Não aplicável", 3

    FormularioDados.cmbPlanejador.Clear
    FormularioDados.cmbPlanejador.AddItem "MAGLIONI", 0
    FormularioDados.cmbPlanejador.AddItem "MAURICIOFA", 1
    FormularioDados.cmbPlanejador.AddItem "CAMILARM", 2
    
    FormularioDados.TabStrip1.Value = 0 'RESETA PARA TAB DO PCP
    
    result = MsgBox("Deseja carregar os dados da ultima execução?", vbYesNo + vbQuestion, "Carregar dados")
    If result = vbYes Then
        Application.EnableEvents = False
        Call set_controlStatus(FormularioDados, "FieldValues")
        Application.EnableEvents = True
    Else
        Call get_controlStatus(FormularioDados, "InitialValues")
    End If

End If

End Sub
```

#### `UserForm_Terminate` (Sub)

```vb
Private Sub UserForm_Terminate()

    Call get_controlStatus(FormularioDados, "FieldValues")
    Application.EnableEvents = True

End Sub
```


## Horas_Orçadas_Salas

### Declarações do módulo (topo do arquivo)

```vb
Attribute VB_Name = "Horas_Orçadas_Salas"
Sub Horas_Orçadas()

    Dim SapGuiAuto As Object
    Dim Application As Object
    Dim Connection As Object
    Dim Session As Object
    Dim WScript As Object
    Dim diagrEngMec As String
    Dim diagrMec1 As String
    Dim diagrEle As String
    
    
    'Conexão com o Objeto SAP
    Set SapGuiAuto = GetObject("SAPGUI")
    Set Application = SapGuiAuto.GetScriptingEngine
    Set Connection = Application.Children(0)
    
    sessao_ok = False
    For j = 0 To Application.Children(0).Sessions.Count() - 1
        Set Session = Connection.Children(CLng(j))
        If Session.ActiveWindow.Text = "Project Builder: subprojeto " & FormularioDados.txtPEP.Value Then
            sessao_ok = True
            Exit For
        End If
    Next
    
    'Detectar qual é o diagrama diagrama da Engenharia
    Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "NETW_OVW"
    nr_max_linhas_1 = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010").verticalScrollbar.Maximum + 1
    For x = 0 To nr_max_linhas_1 - 1
        If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & x & "]").Text = "Engenharia Mecânica" Then
            diagrEngMec = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtNETW_OVW-AUFNR[0," & x & "]").Text
            Exit For
        End If
    Next x
   
    'Detectar qual é o diagrama diagrama da Mecânica
    Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "NETW_OVW"
    nr_max_linhas_1 = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010").verticalScrollbar.Maximum + 1
    For i = 0 To nr_max_linhas_1 - 1
        If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & i & "]").Text = "Mecânica 1" Then
            diagrMec1 = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtNETW_OVW-AUFNR[0," & i & "]").Text
            Exit For
        End If
    Next i
    
    'Detectar qual é o diagrama diagrama da Elétrica
    Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "NETW_OVW"
    nr_max_linhas_1 = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010").verticalScrollbar.Maximum + 1
    For f = 0 To nr_max_linhas_1 - 1
        If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & f & "]").Text = "Elétrica" Then
            diagrEle = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtNETW_OVW-AUFNR[0," & f & "]").Text
            Exit For
        End If
    Next f
    
    Dim HorLOM_OK As Boolean
    Dim HorEDF_OK As Boolean
    Dim HorIST_OK As Boolean
    Dim nr_max_linhas As Integer
    Dim nr_linhas_visiveis As Integer
    
    HorLOM_OK = False
    HorEDF_OK = False
    HorIST_OK = False
    
    'Horas Orçadas para Engenharia Mecânica
    If FormularioDados.txtHorLOM <> "" Then
        Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "NETW_OVW"
        nr_max_linhas_1 = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010").verticalScrollbar.Maximum + 1
        For x = 0 To nr_max_linhas_1 - 1
            If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & x & "]").Text = "Engenharia Mecânica" Then
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtNETW_OVW-AUFNR[0," & x & "]").SetFocus
                Session.findById("wnd[0]").sendVKey 2
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "ACTY_OVW"
                nr_max_linhas = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Maximum + 1
                nr_linhas_visiveis = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").visiblerowcount
                For a = nr_linhas_visiveis - 1 To 0 Step -1
                    If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-VORNR[0," & a & "]").Text = "0500" Then
                        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(a).Selected = True
                        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/btnBUTTON_SELECT_DETAIL").press
                        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCONW:1001/tabsTABSTRIP_1000/tabpUSER").Select
                        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCONW:1001/tabsTABSTRIP_1000/tabpUSER/ssubSUBSCR_1000:SAPLCONW:1320/ctxtAFVGD-USR06").Text = FormularioDados.txtHorLOM.Text
                        Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
                        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "ACTY_OVW"
                        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(a).Selected = False
                        HorLOM_OK = True
                        Exit For
                    End If
                Next
                If HorLOM_OK = False Then
                    If (nr_max_linhas - (nr_linhas_visiveis - 1)) > (2 * nr_linhas_visiveis) Then
                        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = nr_linhas_visiveis * 2
                        y = nr_linhas_visiveis * 3
                    Else
                        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = nr_linhas_visiveis
                        y = nr_linhas_visiveis * 2
                    End If
                    For a = 0 To nr_max_linhas - y
                        If a <> 0 Then
                            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position + 1
                        End If
                        If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-VORNR[0,0]").Text = "0500" Then
                           Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position).Selected = True
                            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/btnBUTTON_SELECT_DETAIL").press
                            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCONW:1001/tabsTABSTRIP_1000/tabpUSER").Select
                            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCONW:1001/tabsTABSTRIP_1000/tabpUSER/ssubSUBSCR_1000:SAPLCONW:1320/ctxtAFVGD-USR06").Text = FormularioDados.txtHorLOM.Text
                            Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
                            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "ACTY_OVW"
                            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position).Selected = False
                            HorLOM_OK = True
                            Exit For
                        End If
                    Next
                End If
            End If
        Exit For
        Next
    End If
    'Horas Orçadas para Edificação
    If FormularioDados.txtHorEDF <> "" Then
        Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "NETW_OVW"
        nr_max_linhas_1 = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010").verticalScrollbar.Maximum + 1
        For x = 0 To nr_max_linhas_1 - 1
            If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & x & "]").Text = "Mecânica 1" Then
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtNETW_OVW-AUFNR[0," & x & "]").SetFocus
                Session.findById("wnd[0]").sendVKey 2
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "ACTY_OVW"
                nr_max_linhas = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Maximum + 1
                nr_linhas_visiveis = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").visiblerowcount
                For a = nr_linhas_visiveis - 1 To 0 Step -1
                    If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-VORNR[0," & a & "]").Text = "0750" Then
                        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(a).Selected = True
                        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/btnBUTTON_SELECT_DETAIL").press
                        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCONW:1001/tabsTABSTRIP_1000/tabpUSER").Select
                        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCONW:1001/tabsTABSTRIP_1000/tabpUSER/ssubSUBSCR_1000:SAPLCONW:1320/ctxtAFVGD-USR06").Text = FormularioDados.txtHorEDF.Text
                        Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
                        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "ACTY_OVW"
                        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(a).Selected = False
                        HorEDF_OK = True
                        Exit For
                    End If
                Next
                If HorEDF_OK = False Then
                    If (nr_max_linhas - (nr_linhas_visiveis - 1)) > (2 * nr_linhas_visiveis) Then
                        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = nr_linhas_visiveis * 2
                        y = nr_linhas_visiveis * 3
                    Else
                        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = nr_linhas_visiveis
                        y = nr_linhas_visiveis * 2
                    End If
                    For a = 0 To nr_max_linhas - y
                        If a <> 0 Then
                            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position + 1
                        End If
                        If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-VORNR[0,0]").Text = "750" Then
                           Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position).Selected = True
                            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/btnBUTTON_SELECT_DETAIL").press
                            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCONW:1001/tabsTABSTRIP_1000/tabpUSER").Select
                            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCONW:1001/tabsTABSTRIP_1000/tabpUSER/ssubSUBSCR_1000:SAPLCONW:1320/ctxtAFVGD-USR06").Text = FormularioDados.txtHorEDF.Text
                            Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
                            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "ACTY_OVW"
                            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position).Selected = False
                            HorEDF_OK = True
                            Exit For
                        End If
                    Next
                End If
            Exit For
            End If
        Next
    End If
    'Horas Orçadas para Elétrica
    If FormularioDados.txtHorINT <> "" Then
        Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "NETW_OVW"
        nr_max_linhas_1 = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010").verticalScrollbar.Maximum + 1
        For x = 0 To nr_max_linhas_1 - 1
            If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & x & "]").Text = "Elétrica" Then
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtNETW_OVW-AUFNR[0," & x & "]").SetFocus
                Session.findById("wnd[0]").sendVKey 2
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "ACTY_OVW"
                nr_max_linhas = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Maximum + 1
                nr_linhas_visiveis = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").visiblerowcount
                For a = nr_linhas_visiveis - 1 To 0 Step -1
                    If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-VORNR[0," & a & "]").Text = "0870" Then
                        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(a).Selected = True
                        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/btnBUTTON_SELECT_DETAIL").press
                        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCONW:1001/tabsTABSTRIP_1000/tabpUSER").Select
                        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCONW:1001/tabsTABSTRIP_1000/tabpUSER/ssubSUBSCR_1000:SAPLCONW:1320/ctxtAFVGD-USR06").Text = FormularioDados.txtHorINT.Text
                        Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
                        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "ACTY_OVW"
                        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(a).Selected = False
                        HorIST_OK = True
                        Exit For
                    End If
                Next
                If HorIST_OK = False Then
                    If (nr_max_linhas - (nr_linhas_visiveis - 1)) > (2 * nr_linhas_visiveis) Then
                        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = nr_linhas_visiveis * 2
                        y = nr_linhas_visiveis * 3
                    Else
                        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = nr_linhas_visiveis
                        y = nr_linhas_visiveis * 2
                    End If
                    For a = 0 To nr_max_linhas - y
                        If a <> 0 Then
                            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position + 1
                        End If
                        If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-VORNR[0,0]").Text = "0870" Then
                           Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position).Selected = True
                            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/btnBUTTON_SELECT_DETAIL").press
                            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCONW:1001/tabsTABSTRIP_1000/tabpUSER").Select
                            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCONW:1001/tabsTABSTRIP_1000/tabpUSER/ssubSUBSCR_1000:SAPLCONW:1320/ctxtAFVGD-USR06").Text = FormularioDados.txtHorINT.Text
                            Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
                            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "ACTY_OVW"
                            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position).Selected = False
                            HorIST_OK = True
                            Exit For
                        End If
                    Next
                End If
            Exit For
            End If
        Next
    End If
End Sub
```


## Alterar_CTs_Solar

### Declarações do módulo (topo do arquivo)

```vb
Attribute VB_Name = "Alterar_CTs_Solar"
```

### Procedimentos

#### `alterar_CTs` (Sub)

```vb
Sub alterar_CTs()

    Dim SapGuiAuto As Object
    Dim Application As Object
    Dim Connection As Object
    Dim Session As Object
    Dim WScript As Object
    
    'Conexão com o Objeto SAP
    Set SapGuiAuto = GetObject("SAPGUI")
    Set Application = SapGuiAuto.GetScriptingEngine
    Set Connection = Application.Children(0)
    
    sessao_ok = False
    For j = 0 To Application.Children(0).Sessions.Count() - 1
        Set Session = Connection.Children(CLng(j))
        If Session.ActiveWindow.Text = "SAP Easy Access" Then
            sessao_ok = True
            Exit For
        End If
    Next
    
    'If sessao_ok = False Then
    '    MsgBox "Nenhuma janela do SAP na tela inicial foi encontrada. Programa interrompido.", vbOKOnly
    '    Exit Sub
    'End If
    
    Dim nr_linhas_visiveis As Integer
    Dim nr_max_linhas As Integer
    Dim x As Integer
    Dim y As Integer
    Dim ALO_ok As Boolean    'ALO - Aprovação Layout
    Dim AAC_ok As Boolean    'AAC - Aprovação Projeto de Acessórios
    Dim LOG_ok As Boolean    'LOG - Logística
    Dim AIL_ok As Boolean    'AIL - Aprovação Proj. Ilumin. e Tomadas
    Dim ACL_ok As Boolean    'ACL - Aprovação Projeto Climatização
    Dim AIN_ok As Boolean    'AIN - Aprovação Projeto Incêndio
    Dim ASS_ok As Boolean    'ASS - Aprovação Proj. Sist. de Segurança
    Dim ADI_ok As Boolean    'ADI - Aprovação Diagrama de Interligação
    Dim ABA_ok As Boolean    'ABA - Aprovação Projeto Bandej./Aterr.
    Dim FAT_ok As Boolean    'FAT - Faturamento
    Dim ENC_ok As Boolean    'ENC - Encerramento
     
  
    Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000003"
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "ACTY_OVW"
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = 0
    nr_max_linhas = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Maximum + 1
    nr_linhas_visiveis = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").visiblerowcount

    'ALO - Aprovação Layout
    ALO_ok = False
    If nr_max_linhas > nr_linhas_visiveis Then
        y = nr_linhas_visiveis - 1
    Else
        y = nr_max_linhas - 1
    End If
    For x = 0 To y
        If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-VORNR[0," & x & "]").Text = "0510" Then
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(x).Selected = True
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/btnBUTTON_SELECT_DETAIL").press
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCONW:1001/tabsTABSTRIP_1000/tabpARBD/ssubSUBSCR_1000:SAPLCONW:1310/ctxtAFVGD-ARBPL").Text = "04030885"
            Session.findById("wnd[0]").sendVKey 0
            Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "ACTY_OVW"
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(x).Selected = False
            ALO_ok = True
            Exit For
        End If
    Next
    Session.findById("wnd[0]").sendVKey 0
    
    'Loop para encontar a tarefa desejada e fazer a substiruíção do centro de trabalho
    If ALO_ok = False Then
        If nr_max_linhas > (2 * nr_linhas_visiveis) Then
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = nr_linhas_visiveis * 1
            y = nr_linhas_visiveis * 2
        Else
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = nr_linhas_visiveis
            y = nr_linhas_visiveis * 1
        End If
        For x = 0 To nr_max_linhas - y
            If x <> 0 Then
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position + 1
            End If
           If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-VORNR[0,0]").Text = "0510" Then
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position).Selected = True
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/btnBUTTON_SELECT_DETAIL").press
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCONW:1001/tabsTABSTRIP_1000/tabpARBD/ssubSUBSCR_1000:SAPLCONW:1310/ctxtAFVGD-ARBPL").Text = "04030885"
                Session.findById("wnd[0]").sendVKey 0
                Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "ACTY_OVW"
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position).Selected = False
                ALO_ok = True
                Exit For
            End If
        Next
        Session.findById("wnd[0]").sendVKey 0
    End If
    
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'AAC - Aprovação Projeto de Acessórios
    AAC_ok = False
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = 0
    If nr_max_linhas > nr_linhas_visiveis Then
        y = nr_linhas_visiveis - 1
    Else
        y = nr_max_linhas - 1
    End If
    For x = 0 To y
        If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-VORNR[0," & x & "]").Text = "0575" Then
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(x).Selected = True
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/btnBUTTON_SELECT_DETAIL").press
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCONW:1001/tabsTABSTRIP_1000/tabpARBD/ssubSUBSCR_1000:SAPLCONW:1310/ctxtAFVGD-ARBPL").Text = "04030885"
            Session.findById("wnd[0]").sendVKey 0
            Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "ACTY_OVW"
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(x).Selected = False
            AAC_ok = True
            Exit For
        End If
    Next
    Session.findById("wnd[0]").sendVKey 0
    
    'Loop para encontar a tarefa desejada e fazer a substiruíção do centro de trabalho
    If AAC_ok = False Then
        If nr_max_linhas > (2 * nr_linhas_visiveis) Then
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = nr_linhas_visiveis * 1
            y = nr_linhas_visiveis * 2
        Else
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = nr_linhas_visiveis
            y = nr_linhas_visiveis * 1
        End If
        For x = 0 To nr_max_linhas - y
            If x <> 0 Then
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position + 1
            End If
           If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-VORNR[0,0]").Text = "0575" Then
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position).Selected = True
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/btnBUTTON_SELECT_DETAIL").press
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCONW:1001/tabsTABSTRIP_1000/tabpARBD/ssubSUBSCR_1000:SAPLCONW:1310/ctxtAFVGD-ARBPL").Text = "04030885"
                Session.findById("wnd[0]").sendVKey 0
                Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "ACTY_OVW"
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position).Selected = False
                AAC_ok = True
                Exit For
            End If
        Next
        Session.findById("wnd[0]").sendVKey 0
    End If
    
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'LOG - Logística
    LOG_ok = False
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = 0
    If nr_max_linhas > nr_linhas_visiveis Then
        y = nr_linhas_visiveis - 1
    Else
        y = nr_max_linhas - 1
    End If
    For x = 0 To y
        If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-VORNR[0," & x & "]").Text = "0598" Then
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(x).Selected = True
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/btnBUTTON_SELECT_DETAIL").press
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCONW:1001/tabsTABSTRIP_1000/tabpARBD/ssubSUBSCR_1000:SAPLCONW:1310/ctxtAFVGD-ARBPL").Text = "04030890"
            Session.findById("wnd[0]").sendVKey 0
            Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "ACTY_OVW"
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(x).Selected = False
            LOG_ok = True
            Exit For
        End If
    Next
    Session.findById("wnd[0]").sendVKey 0
    
    'Loop para encontar a tarefa desejada e fazer a substiruíção do centro de trabalho
    If LOG_ok = False Then
        If nr_max_linhas > (2 * nr_linhas_visiveis) Then
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = nr_linhas_visiveis * 1
            y = nr_linhas_visiveis * 2
        Else
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = nr_linhas_visiveis
            y = nr_linhas_visiveis * 1
        End If
        For x = 0 To nr_max_linhas - y
            If x <> 0 Then
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position + 1
            End If
           If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-VORNR[0,0]").Text = "0598" Then
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position).Selected = True
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/btnBUTTON_SELECT_DETAIL").press
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCONW:1001/tabsTABSTRIP_1000/tabpARBD/ssubSUBSCR_1000:SAPLCONW:1310/ctxtAFVGD-ARBPL").Text = "04030890"
                Session.findById("wnd[0]").sendVKey 0
                Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "ACTY_OVW"
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position).Selected = False
                LOG_ok = True
                Exit For
            End If
        Next
        Session.findById("wnd[0]").sendVKey 0
    End If
    
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'AIL - Aprovação Proj. Ilumin. e Tomadas
    AIL_ok = False
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = 0
    If nr_max_linhas > nr_linhas_visiveis Then
        y = nr_linhas_visiveis - 1
    Else
        y = nr_max_linhas - 1
    End If
    For x = 0 To y
        If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-VORNR[0," & x & "]").Text = "0605" Then
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(x).Selected = True
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/btnBUTTON_SELECT_DETAIL").press
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCONW:1001/tabsTABSTRIP_1000/tabpARBD/ssubSUBSCR_1000:SAPLCONW:1310/ctxtAFVGD-ARBPL").Text = "04030885"
            Session.findById("wnd[0]").sendVKey 0
            Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "ACTY_OVW"
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(x).Selected = False
            AIL_ok = True
            Exit For
        End If
    Next
    Session.findById("wnd[0]").sendVKey 0
    
    'Loop para encontar a tarefa desejada e fazer a substiruíção do centro de trabalho
    If AIL_ok = False Then
        If nr_max_linhas > (2 * nr_linhas_visiveis) Then
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = nr_linhas_visiveis * 1
            y = nr_linhas_visiveis * 2
        Else
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = nr_linhas_visiveis
            y = nr_linhas_visiveis * 1
        End If
        For x = 0 To nr_max_linhas - y
            If x <> 0 Then
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position + 1
            End If
           If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-VORNR[0,0]").Text = "0605" Then
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position).Selected = True
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/btnBUTTON_SELECT_DETAIL").press
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCONW:1001/tabsTABSTRIP_1000/tabpARBD/ssubSUBSCR_1000:SAPLCONW:1310/ctxtAFVGD-ARBPL").Text = "04030885"
                Session.findById("wnd[0]").sendVKey 0
                Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "ACTY_OVW"
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position).Selected = False
                AIL_ok = True
                Exit For
            End If
        Next
        Session.findById("wnd[0]").sendVKey 0
    End If
    
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'ACL - Aprovação Projeto Climatização
    ACL_ok = False
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = 0
    If nr_max_linhas > nr_linhas_visiveis Then
        y = nr_linhas_visiveis - 1
    Else
        y = nr_max_linhas - 1
    End If
    For x = 0 To y
        If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-VORNR[0," & x & "]").Text = "0615" Then
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(x).Selected = True
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/btnBUTTON_SELECT_DETAIL").press
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCONW:1001/tabsTABSTRIP_1000/tabpARBD/ssubSUBSCR_1000:SAPLCONW:1310/ctxtAFVGD-ARBPL").Text = "04030885"
            Session.findById("wnd[0]").sendVKey 0
            Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "ACTY_OVW"
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(x).Selected = False
            ACL_ok = True
            Exit For
        End If
    Next
    Session.findById("wnd[0]").sendVKey 0
    
    'Loop para encontar a tarefa desejada e fazer a substiruíção do centro de trabalho
    If ACL_ok = False Then
        If nr_max_linhas > (2 * nr_linhas_visiveis) Then
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = nr_linhas_visiveis * 1
            y = nr_linhas_visiveis * 2
        Else
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = nr_linhas_visiveis
            y = nr_linhas_visiveis * 1
        End If
        For x = 0 To nr_max_linhas - y
            If x <> 0 Then
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position + 1
            End If
           If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-VORNR[0,0]").Text = "0615" Then
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position).Selected = True
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/btnBUTTON_SELECT_DETAIL").press
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCONW:1001/tabsTABSTRIP_1000/tabpARBD/ssubSUBSCR_1000:SAPLCONW:1310/ctxtAFVGD-ARBPL").Text = "04030885"
                Session.findById("wnd[0]").sendVKey 0
                Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "ACTY_OVW"
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position).Selected = False
                ACL_ok = True
                Exit For
            End If
        Next
        Session.findById("wnd[0]").sendVKey 0
    End If
    
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'AIN - Aprovação Projeto Incêndio
    AIN_ok = False
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = 0
    If nr_max_linhas > nr_linhas_visiveis Then
        y = nr_linhas_visiveis - 1
    Else
        y = nr_max_linhas - 1
    End If
    For x = 0 To y
        If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-VORNR[0," & x & "]").Text = "0635" Then
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(x).Selected = True
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/btnBUTTON_SELECT_DETAIL").press
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCONW:1001/tabsTABSTRIP_1000/tabpARBD/ssubSUBSCR_1000:SAPLCONW:1310/ctxtAFVGD-ARBPL").Text = "04030885"
            Session.findById("wnd[0]").sendVKey 0
            Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "ACTY_OVW"
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(x).Selected = False
            AIN_ok = True
            Exit For
        End If
    Next
    Session.findById("wnd[0]").sendVKey 0
    
    'Loop para encontar a tarefa desejada e fazer a substiruíção do centro de trabalho
    If AIN_ok = False Then
        If nr_max_linhas > (2 * nr_linhas_visiveis) Then
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = nr_linhas_visiveis * 1
            y = nr_linhas_visiveis * 2
        Else
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = nr_linhas_visiveis
            y = nr_linhas_visiveis * 1
        End If
        For x = 0 To nr_max_linhas - y
            If x <> 0 Then
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position + 1
            End If
           If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-VORNR[0,0]").Text = "0635" Then
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position).Selected = True
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/btnBUTTON_SELECT_DETAIL").press
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCONW:1001/tabsTABSTRIP_1000/tabpARBD/ssubSUBSCR_1000:SAPLCONW:1310/ctxtAFVGD-ARBPL").Text = "04030885"
                Session.findById("wnd[0]").sendVKey 0
                Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "ACTY_OVW"
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position).Selected = False
                AIN_ok = True
                Exit For
            End If
        Next
        Session.findById("wnd[0]").sendVKey 0
    End If
    
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'ASS - Aprovação Proj. Sist. de Segurança
    ASS_ok = False
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = 0
    If nr_max_linhas > nr_linhas_visiveis Then
        y = nr_linhas_visiveis - 1
    Else
        y = nr_max_linhas - 1
    End If
    For x = 0 To y
        If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-VORNR[0," & x & "]").Text = "0655" Then
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(x).Selected = True
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/btnBUTTON_SELECT_DETAIL").press
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCONW:1001/tabsTABSTRIP_1000/tabpARBD/ssubSUBSCR_1000:SAPLCONW:1310/ctxtAFVGD-ARBPL").Text = "04030885"
            Session.findById("wnd[0]").sendVKey 0
            Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "ACTY_OVW"
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(x).Selected = False
            ASS_ok = True
            Exit For
        End If
    Next
    Session.findById("wnd[0]").sendVKey 0
    
    'Loop para encontar a tarefa desejada e fazer a substiruíção do centro de trabalho
    If ASS_ok = False Then
        If nr_max_linhas > (2 * nr_linhas_visiveis) Then
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = nr_linhas_visiveis * 1
            y = nr_linhas_visiveis * 2
        Else
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = nr_linhas_visiveis
            y = nr_linhas_visiveis * 1
        End If
        For x = 0 To nr_max_linhas - y
            If x <> 0 Then
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position + 1
            End If
           If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-VORNR[0,0]").Text = "0655" Then
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position).Selected = True
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/btnBUTTON_SELECT_DETAIL").press
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCONW:1001/tabsTABSTRIP_1000/tabpARBD/ssubSUBSCR_1000:SAPLCONW:1310/ctxtAFVGD-ARBPL").Text = "04030885"
                Session.findById("wnd[0]").sendVKey 0
                Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "ACTY_OVW"
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position).Selected = False
                ASS_ok = True
                Exit For
            End If
        Next
        Session.findById("wnd[0]").sendVKey 0
    End If
    
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'ADI - Aprovação Diagrama de Interligação
    ADI_ok = False
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = 0
    If nr_max_linhas > nr_linhas_visiveis Then
        y = nr_linhas_visiveis - 1
    Else
        y = nr_max_linhas - 1
    End If
    For x = 0 To y
        If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-VORNR[0," & x & "]").Text = "0685" Then
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(x).Selected = True
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/btnBUTTON_SELECT_DETAIL").press
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCONW:1001/tabsTABSTRIP_1000/tabpARBD/ssubSUBSCR_1000:SAPLCONW:1310/ctxtAFVGD-ARBPL").Text = "04030885"
            Session.findById("wnd[0]").sendVKey 0
            Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "ACTY_OVW"
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(x).Selected = False
            ADI_ok = True
            Exit For
        End If
    Next
    Session.findById("wnd[0]").sendVKey 0
    
    'Loop para encontar a tarefa desejada e fazer a substiruíção do centro de trabalho
    If ADI_ok = False Then
        If nr_max_linhas > (2 * nr_linhas_visiveis) Then
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = nr_linhas_visiveis * 1
            y = nr_linhas_visiveis * 2
        Else
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = nr_linhas_visiveis
            y = nr_linhas_visiveis * 1
        End If
        For x = 0 To nr_max_linhas - y
            If x <> 0 Then
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position + 1
            End If
           If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-VORNR[0,0]").Text = "0685" Then
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position).Selected = True
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/btnBUTTON_SELECT_DETAIL").press
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCONW:1001/tabsTABSTRIP_1000/tabpARBD/ssubSUBSCR_1000:SAPLCONW:1310/ctxtAFVGD-ARBPL").Text = "04030885"
                Session.findById("wnd[0]").sendVKey 0
                Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "ACTY_OVW"
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position).Selected = False
                ADI_ok = True
                Exit For
            End If
        Next
        Session.findById("wnd[0]").sendVKey 0
    End If
    
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'ABA - Aprovação Projeto Bandej./Aterr.
    ABA_ok = False
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = 0
    If nr_max_linhas > nr_linhas_visiveis Then
        y = nr_linhas_visiveis - 1
    Else
        y = nr_max_linhas - 1
    End If
    For x = 0 To y
        If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-VORNR[0," & x & "]").Text = "0695" Then
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(x).Selected = True
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/btnBUTTON_SELECT_DETAIL").press
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCONW:1001/tabsTABSTRIP_1000/tabpARBD/ssubSUBSCR_1000:SAPLCONW:1310/ctxtAFVGD-ARBPL").Text = "04030885"
            Session.findById("wnd[0]").sendVKey 0
            Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "ACTY_OVW"
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(x).Selected = False
            ABA_ok = True
            Exit For
        End If
    Next
    Session.findById("wnd[0]").sendVKey 0
    
    'Loop para encontar a tarefa desejada e fazer a substiruíção do centro de trabalho
    If ABA_ok = False Then
        If nr_max_linhas > (2 * nr_linhas_visiveis) Then
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = nr_linhas_visiveis * 1
            y = nr_linhas_visiveis * 2
        Else
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = nr_linhas_visiveis
            y = nr_linhas_visiveis * 1
        End If
        For x = 0 To nr_max_linhas - y
            If x <> 0 Then
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position + 1
            End If
           If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-VORNR[0,0]").Text = "0695" Then
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position).Selected = True
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/btnBUTTON_SELECT_DETAIL").press
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCONW:1001/tabsTABSTRIP_1000/tabpARBD/ssubSUBSCR_1000:SAPLCONW:1310/ctxtAFVGD-ARBPL").Text = "04030885"
                Session.findById("wnd[0]").sendVKey 0
                Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "ACTY_OVW"
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position).Selected = False
                ABA_ok = True
                Exit For
            End If
        Next
        Session.findById("wnd[0]").sendVKey 0
    End If
    
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'FAT - Faturamento.
    FAT_ok = False
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = 0
    If nr_max_linhas > nr_linhas_visiveis Then
        y = nr_linhas_visiveis - 1
    Else
        y = nr_max_linhas - 1
    End If
    For x = 0 To y
        If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-VORNR[0," & x & "]").Text = "0910" Then
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(x).Selected = True
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/btnBUTTON_SELECT_DETAIL").press
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCONW:1001/tabsTABSTRIP_1000/tabpARBD/ssubSUBSCR_1000:SAPLCONW:1310/ctxtAFVGD-ARBPL").Text = "04030830"
            Session.findById("wnd[0]").sendVKey 0
            Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "ACTY_OVW"
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(x).Selected = False
            FAT_ok = True
            Exit For
        End If
    Next
    Session.findById("wnd[0]").sendVKey 0
    
    'Loop para encontar a tarefa desejada e fazer a substiruíção do centro de trabalho
    If FAT_ok = False Then
        If nr_max_linhas > (2 * nr_linhas_visiveis) Then
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = nr_linhas_visiveis * 1
            y = nr_linhas_visiveis * 2
        Else
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = nr_linhas_visiveis
            y = nr_linhas_visiveis * 1
        End If
        For x = 0 To nr_max_linhas - y
            If x <> 0 Then
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position + 1
            End If
           If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-VORNR[0,0]").Text = "0910" Then
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position).Selected = True
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/btnBUTTON_SELECT_DETAIL").press
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCONW:1001/tabsTABSTRIP_1000/tabpARBD/ssubSUBSCR_1000:SAPLCONW:1310/ctxtAFVGD-ARBPL").Text = "04030830"
                Session.findById("wnd[0]").sendVKey 0
                Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "ACTY_OVW"
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position).Selected = False
                FAT_ok = True
                Exit For
            End If
        Next
        Session.findById("wnd[0]").sendVKey 0
    End If
    
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'ENC - Encerramento.
    ENC_ok = False
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = 0
    If nr_max_linhas > nr_linhas_visiveis Then
        y = nr_linhas_visiveis - 1
    Else
        y = nr_max_linhas - 1
    End If
    For x = 0 To y
        If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-VORNR[0," & x & "]").Text = "0920" Then
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(x).Selected = True
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/btnBUTTON_SELECT_DETAIL").press
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCONW:1001/tabsTABSTRIP_1000/tabpARBD/ssubSUBSCR_1000:SAPLCONW:1310/ctxtAFVGD-ARBPL").Text = "04030890"
            Session.findById("wnd[0]").sendVKey 0
            Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "ACTY_OVW"
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(x).Selected = False
            ENC_ok = True
            Exit For
        End If
    Next
    Session.findById("wnd[0]").sendVKey 0
    
    'Loop para encontar a tarefa desejada e fazer a substiruíção do centro de trabalho
    If ENC_ok = False Then
        If nr_max_linhas > (2 * nr_linhas_visiveis) Then
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = nr_linhas_visiveis * 1
            y = nr_linhas_visiveis * 2
        Else
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = nr_linhas_visiveis
            y = nr_linhas_visiveis * 1
        End If
        For x = 0 To nr_max_linhas - y
            If x <> 0 Then
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position + 1
            End If
           If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-VORNR[0,0]").Text = "0920" Then
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position).Selected = True
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/btnBUTTON_SELECT_DETAIL").press
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCONW:1001/tabsTABSTRIP_1000/tabpARBD/ssubSUBSCR_1000:SAPLCONW:1310/ctxtAFVGD-ARBPL").Text = "04030890"
                Session.findById("wnd[0]").sendVKey 0
                Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "ACTY_OVW"
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position).Selected = False
                ENC_ok = True
                Exit For
            End If
        Next
        Session.findById("wnd[0]").sendVKey 0
    End If
   
    
End Sub
```


## Plan_Tempos

### Declarações do módulo (topo do arquivo)

```vb
Attribute VB_Name = "Plan_Tempos"
Public erro_DR As Boolean
Public HorLOM As Double
Public DurLOM As Integer
Public HorLMM As Double
Public DurLMM As Integer
Public HorPBS As Double
Public DurPBS As Integer
Public HorPPA As Double
Public DurPPA As Integer
Public HorPCI As Double
Public DurPCI As Integer
Public HorPCE As Double
Public DurPCE As Integer
Public HorPAC As Double
Public DurPAC As Integer
Public HorLCA As Double
Public DurLCA As Integer
Public HorLAA As Double
Public DurLAA As Integer
Public HorLAM As Double
Public DurLAM As Integer
Public HorLMA As Double
Public DurLMA As Integer
Public HorPTR As Double
Public DurPTR As Integer

Public Hor531 As Double
Public Dur531 As Double
Public Hor551 As Double
Public Dur551 As Double
Public Hor561 As Double
Public Dur561 As Double
Public Hor581 As Double
Public Dur581 As Double
Public Hor585 As Double
Public Dur585 As Double
Public Hor589 As Double
Public Dur589 As Double
Public modulos As Integer

Public HorPIL As Double
Public DurPIL As Integer
Public HorPCL As Double
Public DurPCL As Integer
Public HorCSM As Double
Public DurCSM As Integer
Public HorLMC As Double
Public DurLMC As Integer
Public HorPIN As Double
Public DurPIN As Integer
Public HorLMI As Double
Public DurLMI As Integer
Public HorPSS As Double
Public DurPSS As Integer
Public HorLMS As Double
Public DurLMS As Integer
Public HorPBA As Double
Public DurPBA As Integer
Public HorDIN As Double
Public DurDIN As Integer
Public HorLMT As Double
Public DurLMT As Integer
Public HorLBA As Double
Public DurLBA As Integer
Public HorLMD As Double
Public DurLMD As Integer
Public HorPRF As Double
Public DurPRF As Integer

Public HorCOR1 As Double
Public DurCOR1 As Integer
Public HorFCH1 As Double
Public DurFCH1 As Integer
Public HorPRB1 As Double
Public DurPRB1 As Integer
Public HorSBA1 As Double
Public DurSBA1 As Integer
Public HorPRE1 As Double
Public DurPRE1 As Integer
Public HorSES1 As Double
Public DurSES1 As Integer
Public HorEDF1 As Double
Public DurEDF1 As Integer
Public HorPIN1 As Double
Public DurPIN1 As Integer
Public HorCHI1 As Double
Public DurCHI1 As Integer
Public HorCHE1 As Double
Public DurCHE1 As Integer

Public HorCOR2 As Double
Public DurCOR2 As Integer
Public HorFCH2 As Double
Public DurFCH2 As Integer
Public HorPRB2 As Double
Public DurPRB2 As Integer
Public HorSBA2 As Double
Public DurSBA2 As Integer
Public HorPRE2 As Double
Public DurPRE2 As Integer
Public HorSES2 As Double
Public DurSES2 As Integer
Public HorEDF2 As Double
Public DurEDF2 As Integer
Public HorPIN2 As Double
Public DurPIN2 As Integer
Public HorCHI2 As Double
Public DurCHI2 As Integer
Public HorCHE2 As Double
Public DurCHE2 As Integer

Public HorCOR3 As Double
Public DurCOR3 As Integer
Public HorFCH3 As Double
Public DurFCH3 As Integer
Public HorPRB3 As Double
Public DurPRB3 As Integer
Public HorSBA3 As Double
Public DurSBA3 As Integer
Public HorPRE3 As Double
Public DurPRE3 As Integer
Public HorSES3 As Double
Public DurSES3 As Integer
Public HorEDF3 As Double
Public DurEDF3 As Integer
Public HorPIN3 As Double
Public DurPIN3 As Integer
Public HorCHI3 As Double
Public DurCHI3 As Integer
Public HorCHE3 As Double
Public DurCHE3 As Integer

Public HorCOR4 As Double
Public DurCOR4 As Integer
Public HorFCH4 As Double
Public DurFCH4 As Integer
Public HorPRB4 As Double
Public DurPRB4 As Integer
Public HorSBA4 As Double
Public DurSBA4 As Integer
Public HorPRE4 As Double
Public DurPRE4 As Integer
Public HorSES4 As Double
Public DurSES4 As Integer
Public HorEDF4 As Double
Public DurEDF4 As Integer
Public HorPIN4 As Double
Public DurPIN4 As Integer
Public HorCHI4 As Double
Public DurCHI4 As Integer
Public HorCHE4 As Double
Public DurCHE4 As Integer

Public HorCOR5 As Double
Public DurCOR5 As Integer
Public HorFCH5 As Double
Public DurFCH5 As Integer
Public HorPRB5 As Double
Public DurPRB5 As Integer
Public HorSBA5 As Double
Public DurSBA5 As Integer
Public HorPRE5 As Double
Public DurPRE5 As Integer
Public HorSES5 As Double
Public DurSES5 As Integer
Public HorEDF5 As Double
Public DurEDF5 As Integer
Public HorPIN5 As Double
Public DurPIN5 As Integer
Public HorCHI5 As Double
Public DurCHI5 As Integer
Public HorCHE5 As Double
Public DurCHE5 As Integer

Public HorCOR6 As Double
Public DurCOR6 As Integer
Public HorFCH6 As Double
Public DurFCH6 As Integer
Public HorPRB6 As Double
Public DurPRB6 As Integer
Public HorSBA6 As Double
Public DurSBA6 As Integer
Public HorPRE6 As Double
Public DurPRE6 As Integer
Public HorSES6 As Double
Public DurSES6 As Integer
Public HorEDF6 As Double
Public DurEDF6 As Integer
Public HorPIN6 As Double
Public DurPIN6 As Integer
Public HorCHI6 As Double
Public DurCHI6 As Integer
Public HorCHE6 As Double
Public DurCHE6 As Integer

Public HorCOR7 As Double
Public DurCOR7 As Integer
Public HorFCH7 As Double
Public DurFCH7 As Integer
Public HorPRB7 As Double
Public DurPRB7 As Integer
Public HorSBA7 As Double
Public DurSBA7 As Integer
Public HorPRE7 As Double
Public DurPRE7 As Integer
Public HorSES7 As Double
Public DurSES7 As Integer
Public HorEDF7 As Double
Public DurEDF7 As Integer
Public HorPIN7 As Double
Public DurPIN7 As Integer
Public HorCHI7 As Double
Public DurCHI7 As Integer
Public HorCHE7 As Double
Public DurCHE7 As Integer

Public HorCOR8 As Double
Public DurCOR8 As Integer
Public HorFCH8 As Double
Public DurFCH8 As Integer
Public HorPRB8 As Double
Public DurPRB8 As Integer
Public HorSBA8 As Double
Public DurSBA8 As Integer
Public HorPRE8 As Double
Public DurPRE8 As Integer
Public HorSES8 As Double
Public DurSES8 As Integer
Public HorEDF8 As Double
Public DurEDF8 As Integer
Public HorPIN8 As Double
Public DurPIN8 As Integer
Public HorCHI8 As Double
Public DurCHI8 As Integer
Public HorCHE8 As Double
Public DurCHE8 As Integer

Public HorFAC As Double
Public DurFAC As Integer
Public HorFCA As Double
Public DurFCA As Integer
Public HorMAM As Double
Public DurMAM As Integer
Public HorMAA As Double
Public DurMAA As Integer

Public HorPRM As Double
Public DurPRM As Integer
Public HorIST As Double
Public DurIST As Integer
Public HorMCL As Double
Public DurMCL As Integer
Public HorMCM As Double
Public DurMCM As Integer
Public HorMIN As Double
Public DurMIN As Integer
Public HorMSS As Double
Public DurMSS As Integer
Public HorFEQ As Double
Public DurFEQ As Integer
Public HorLMB As Double
Public DurLMB As Integer
Public HorINT As Double
Public DurINT As Integer
Public HorTES As Double
Public DurTES As Integer
Public HorINS As Double
Public DurINS As Integer
Public HorPEE As Double
Public DurPEE As Integer
Public HorPEM As Double
Public DurPEM As Integer
Public HorFEC As Double
Public DurFEC As Integer
Public HorFEA As Double
Public DurFEA As Integer
```

### Procedimentos

#### `inicio` (Sub)

```vb
Sub inicio()

Application.ScreenUpdating = False

FormularioDados.Show

Application.ScreenUpdating = True

End Sub
```

#### `calcula` (Sub)

```vb
Sub calcula()

Application.ScreenUpdating = False

If FormularioDados.txtPEP.Value <> "" Then
    ThisWorkbook.Application.StatusBar = "PEP " & FormularioDados.txtPEP.Value & ": Calculando tempos"
Else
    ThisWorkbook.Application.StatusBar = "Calculando tempos"
End If

ActiveSheet.Unprotect

'Limpa aba de resultados
Rows("4:245").Select
Selection.Delete Shift:=xlUp

'Insere template completo
template.Activate

  
If FormularioDados.tipoestrutura.Value = "Container Solar" Or FormularioDados.tipoestrutura.Value = "ESSW (mecânica)" Then
    ActiveSheet.Unprotect
    Range("A1:E153").Select
    Selection.Copy
Else
    ActiveSheet.Unprotect
    Range("G1:K243").Select
    Selection.Copy
End If

resultado.Activate
Range("A3").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False

template.Activate
ActiveSheet.Protect

If FormularioDados.tipoestrutura.Value = "Container Solar" Or FormularioDados.tipoestrutura.Value = "ESSW (mecânica)" Then
    'Formata tabela de tarefas
    resultado.Activate
    Range("A4:E155").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    Range("C4:E155").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("A4:A155").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.NumberFormat = "0000"
    Range("E4:E155").Select
    Selection.NumberFormat = "0.0"
    Range("C4:C155").Select
    Selection.NumberFormat = "0.0"
    
    'Altera nome das tarefas se for Container Solar
    If FormularioDados.tipoestrutura.Value = "Container Solar" Or FormularioDados.tipoestrutura.Value = "ESSW (mecânica)" Then
        Range("B10").Value = "PEC - Projeto Estrutura Container"
        Range("B12").Value = "EMC - Estagiamento Mat. Estrut. Cont."
        Range("B16").Value = "PEI - Projeto Estrutura Interna"
        Range("B18").Value = "EMI - Estagiamento Mat. Estr. Interna"
        Range("B69").Value = "FPC - Fabricação Peças Caldeiraria"
        Range("B74").Value = "OEE - Ordens Estrutura Interna"
        Range("B77").Value = "SEI - Separação Estrutura Interna"
        Range("B78").Value = "MEI - Montagem Estrutura Interna"
        Range("B121").Value = "OFE - Ordens Fechamento Externo"
        Range("B124").Value = "SAF - Separação Almox. Fech. Externo"
        Range("B126").Value = "SFE - Separação Fechamento Externo"
        Range("B127").Value = "MFE - Montagem Fechamento Externo"
    End If
    If FormularioDados.chkFilho.Value = True Then
        Range("A153").Value = 899
    End If
    
    'Altera nome das tarefas de montagem incêndio e leito se necessário
    Range("B131").Value = "SII - Separação Almox. Instal./ Inc."
    Range("B133").Value = "MII - Montagem Instalações / Incêndio"
    If FormularioDados.seguranca.Value <> "Não possui" Then Range("B142").Value = "LBS - Leito e Bandejamento / Sist. Seg."
Else
    'Formata tabela de tarefas
    resultado.Activate
    Range("A4:E245").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    Range("C4:E245").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("A4:A245").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.NumberFormat = "0000"
    Range("E4:E245").Select
    Selection.NumberFormat = "0.0"
    Range("C4:C245").Select
    Selection.NumberFormat = "0.0"
    If FormularioDados.chkFilho.Value = True Then
        Range("A243").Value = 899
    End If
    Range("B221").Value = "SII - Separação Almox. Instal./ Inc."
    Range("B223").Value = "MII - Montagem Instalações / Incêndio"
    If FormularioDados.seguranca.Value <> "Não possui" Then Range("B232").Value = "LBS - Leito e Bandejamento / Sist. Seg."
End If


'###Busca dados de tempos na tabela e preenche no resultado###'
tempos.Activate
ActiveSheet.Unprotect

'Procura linha correspondente
For contador = 4 To 13
    If Range("A" & contador).Value = FormularioDados.nrmodulos.Value Then
        If FormularioDados.tipoestrutura.Value = "Container Solar" Then
            contador = 12
        End If
        If FormularioDados.tipoestrutura.Value = "ESSW (mecânica)" Then
            contador = 13
        End If
        
        If FormularioDados.seguranca.Value <> "Não possui" Then
            Range("B" & contador).Value = "Sim"
        Else
            Range("B" & contador).Value = "Não"
        End If
        
'        If FormularioDados.dutos.Value = True Then
'            Range("C" & contador).Value = "Sim"
'        Else
'            Range("C" & contador).Value = "Não"
'        End If
        
        If FormularioDados.incendio.Value = "Com combate" Then
            Range("D" & contador).Value = "Sim"
        Else
            Range("D" & contador).Value = "Não"
        End If
        
        If FormularioDados.tipomaq.Value = "Roof Top" Then
            Range("E" & contador).Value = "Sim"
        Else
            Range("E" & contador).Value = "Não"
        End If
        
        If FormularioDados.testesw.Value = True Then
            Range("F" & contador).Value = "Sim"
        Else
            Range("F" & contador).Value = "Não"
        End If
        
        If FormularioDados.chaparemov.Value = True Then
            Range("G" & contador).Value = "Sim"
        Else
            Range("G" & contador).Value = "Não"
        End If
        
        If FormularioDados.trafooleo.Value = True Then
            Range("H" & contador).Value = "Sim"
        Else
            Range("H" & contador).Value = "Não"
        End If
        
        If FormularioDados.chkPeDireito = True Then
            Range("I" & contador).Value = "Sim"
        Else
            Range("I" & contador).Value = "Não"
        End If
        
        If FormularioDados.chkEscadaPadrao = True Then
            Range("J" & contador).Value = "Sim"
        Else
            Range("J" & contador).Value = "Não"
        End If
        
        If FormularioDados.chkEscadaEspecial = True Then
            Range("K" & contador).Value = "Sim"
        Else
            Range("K" & contador).Value = "Não"
        End If
        
        If FormularioDados.chkPoraoCabos = True Then
            Range("L" & contador).Value = "Sim"
        Else
            Range("L" & contador).Value = "Não"
        End If
        
        If FormularioDados.chkPilotis = True Then
            Range("M" & contador).Value = "Sim"
        Else
            Range("M" & contador).Value = "Não"
        End If
        
        If FormularioDados.chkRedeDutos = True Then
            Range("N" & contador).Value = "Sim"
        Else
            Range("N" & contador).Value = "Não"
        End If
        
        If FormularioDados.chkFundoFalso = True Then
            Range("O" & contador).Value = "Sim"
        Else
            Range("O" & contador).Value = "Não"
        End If
        
        If FormularioDados.chkDutosBWW = True Then
            Range("P" & contador).Value = "Sim"
        Else
            Range("P" & contador).Value = "Não"
        End If
        
        If FormularioDados.chkCalhasPluviais = True Then
            Range("Q" & contador).Value = "Sim"
        Else
            Range("Q" & contador).Value = "Não"
        End If
    
        If FormularioDados.chk_dutoGases = True Then
            Range("R" & contador).Value = "Sim"
        Else
            Range("R" & contador).Value = "Não"
        End If
        
'        If FormularioDados.chkBaciaContencao = True Then
'            Range("R" & contador).Value = "Sim"
'        Else
'            Range("R" & contador).Value = "Não"
'        End If
'
'        If FormularioDados.chkPisoRemovivel = True Then
'            Range("S" & contador).Value = "Sim"
'        Else
'            Range("S" & contador).Value = "Não"
'        End If
        
        If FormularioDados.nrmodulos.Value = "1 Módulo" And FormularioDados.tipoestrutura.Value <> "Container Solar" And FormularioDados.tipoestrutura.Value <> "ESSW (mecânica)" Then
            Range("BT" & contador).Value = FormularioDados.tipoestrutura.Value
            Range("BU" & contador).Value = CDbl(FormularioDados.modulo1.Value)
            Range("BV" & contador).Value = CDbl(FormularioDados.largmodulo1.Value)
            Range("BW" & contador).Value = FormularioDados.planpin.Value
        End If
        
        If FormularioDados.nrmodulos.Value = "2 Módulos" Then
            Range("BT" & contador).Value = FormularioDados.tipoestrutura.Value
            Range("BU" & contador).Value = CDbl(FormularioDados.modulo1.Value)
            Range("BV" & contador).Value = CDbl(FormularioDados.largmodulo1.Value)
            Range("BW" & contador).Value = FormularioDados.planpin.Value
            
            Range("CR" & contador).Value = FormularioDados.tipoestrutura.Value
            Range("CS" & contador).Value = CDbl(FormularioDados.modulo2.Value)
            Range("CT" & contador).Value = CDbl(FormularioDados.largmodulo2.Value)
            Range("CU" & contador).Value = FormularioDados.planpin.Value
        End If
        
        If FormularioDados.nrmodulos.Value = "3 Módulos" Then
            Range("BT" & contador).Value = FormularioDados.tipoestrutura.Value
            Range("BU" & contador).Value = CDbl(FormularioDados.modulo1.Value)
            Range("BV" & contador).Value = CDbl(FormularioDados.largmodulo1.Value)
            Range("BW" & contador).Value = FormularioDados.planpin.Value
            
            Range("CR" & contador).Value = FormularioDados.tipoestrutura.Value
            Range("CS" & contador).Value = CDbl(FormularioDados.modulo2.Value)
            Range("CT" & contador).Value = CDbl(FormularioDados.largmodulo2.Value)
            Range("CU" & contador).Value = FormularioDados.planpin.Value
            
            Range("DP" & contador).Value = FormularioDados.tipoestrutura.Value
            Range("DQ" & contador).Value = CDbl(FormularioDados.modulo3.Value)
            Range("DR" & contador).Value = CDbl(FormularioDados.largmodulo3.Value)
            Range("DS" & contador).Value = FormularioDados.planpin.Value
        End If
        
        If FormularioDados.nrmodulos.Value = "4 Módulos" Then
            Range("BT" & contador).Value = FormularioDados.tipoestrutura.Value
            Range("BU" & contador).Value = CDbl(FormularioDados.modulo1.Value)
            Range("BV" & contador).Value = CDbl(FormularioDados.largmodulo1.Value)
            Range("BW" & contador).Value = FormularioDados.planpin.Value
            
            Range("CR" & contador).Value = FormularioDados.tipoestrutura.Value
            Range("CS" & contador).Value = CDbl(FormularioDados.modulo2.Value)
            Range("CT" & contador).Value = CDbl(FormularioDados.largmodulo2.Value)
            Range("CU" & contador).Value = FormularioDados.planpin.Value
            
            Range("DP" & contador).Value = FormularioDados.tipoestrutura.Value
            Range("DQ" & contador).Value = CDbl(FormularioDados.modulo3.Value)
            Range("DR" & contador).Value = CDbl(FormularioDados.largmodulo3.Value)
            Range("DS" & contador).Value = FormularioDados.planpin.Value
            
            Range("EN" & contador).Value = FormularioDados.tipoestrutura.Value
            Range("EO" & contador).Value = CDbl(FormularioDados.modulo4.Value)
            Range("EP" & contador).Value = CDbl(FormularioDados.largmodulo4.Value)
            Range("EQ" & contador).Value = FormularioDados.planpin.Value
        End If
        
        If FormularioDados.nrmodulos.Value = "5 Módulos" Then
            Range("BT" & contador).Value = FormularioDados.tipoestrutura.Value
            Range("BU" & contador).Value = CDbl(FormularioDados.modulo1.Value)
            Range("BV" & contador).Value = CDbl(FormularioDados.largmodulo1.Value)
            Range("BW" & contador).Value = FormularioDados.planpin.Value
            
            Range("CR" & contador).Value = FormularioDados.tipoestrutura.Value
            Range("CS" & contador).Value = CDbl(FormularioDados.modulo2.Value)
            Range("CT" & contador).Value = CDbl(FormularioDados.largmodulo2.Value)
            Range("CU" & contador).Value = FormularioDados.planpin.Value
            
            Range("DP" & contador).Value = FormularioDados.tipoestrutura.Value
            Range("DQ" & contador).Value = CDbl(FormularioDados.modulo3.Value)
            Range("DR" & contador).Value = CDbl(FormularioDados.largmodulo3.Value)
            Range("DS" & contador).Value = FormularioDados.planpin.Value
            
            Range("EN" & contador).Value = FormularioDados.tipoestrutura.Value
            Range("EO" & contador).Value = CDbl(FormularioDados.modulo4.Value)
            Range("EP" & contador).Value = CDbl(FormularioDados.largmodulo4.Value)
            Range("EQ" & contador).Value = FormularioDados.planpin.Value
        
            Range("FL" & contador).Value = FormularioDados.tipoestrutura.Value
            Range("FM" & contador).Value = CDbl(FormularioDados.modulo4.Value)
            Range("FN" & contador).Value = CDbl(FormularioDados.largmodulo4.Value)
            Range("FO" & contador).Value = FormularioDados.planpin.Value
        End If
        
        If FormularioDados.nrmodulos.Value = "6 Módulos" Then
            Range("BT" & contador).Value = FormularioDados.tipoestrutura.Value
            Range("BU" & contador).Value = CDbl(FormularioDados.modulo1.Value)
            Range("BV" & contador).Value = CDbl(FormularioDados.largmodulo1.Value)
            Range("BW" & contador).Value = FormularioDados.planpin.Value
            
            Range("CR" & contador).Value = FormularioDados.tipoestrutura.Value
            Range("CS" & contador).Value = CDbl(FormularioDados.modulo2.Value)
            Range("CT" & contador).Value = CDbl(FormularioDados.largmodulo2.Value)
            Range("CU" & contador).Value = FormularioDados.planpin.Value
            
            Range("DP" & contador).Value = FormularioDados.tipoestrutura.Value
            Range("DQ" & contador).Value = CDbl(FormularioDados.modulo3.Value)
            Range("DR" & contador).Value = CDbl(FormularioDados.largmodulo3.Value)
            Range("DS" & contador).Value = FormularioDados.planpin.Value
            
            Range("EN" & contador).Value = FormularioDados.tipoestrutura.Value
            Range("EO" & contador).Value = CDbl(FormularioDados.modulo4.Value)
            Range("EP" & contador).Value = CDbl(FormularioDados.largmodulo4.Value)
            Range("EQ" & contador).Value = FormularioDados.planpin.Value
        
            Range("FL" & contador).Value = FormularioDados.tipoestrutura.Value
            Range("FM" & contador).Value = CDbl(FormularioDados.modulo4.Value)
            Range("FN" & contador).Value = CDbl(FormularioDados.largmodulo4.Value)
            Range("FO" & contador).Value = FormularioDados.planpin.Value
            
            Range("GJ" & contador).Value = FormularioDados.tipoestrutura.Value
            Range("GK" & contador).Value = CDbl(FormularioDados.modulo4.Value)
            Range("GL" & contador).Value = CDbl(FormularioDados.largmodulo4.Value)
            Range("GM" & contador).Value = FormularioDados.planpin.Value
        End If
        
        If FormularioDados.nrmodulos.Value = "7 Módulos" Then
            Range("BT" & contador).Value = FormularioDados.tipoestrutura.Value
            Range("BU" & contador).Value = CDbl(FormularioDados.modulo1.Value)
            Range("BV" & contador).Value = CDbl(FormularioDados.largmodulo1.Value)
            Range("BW" & contador).Value = FormularioDados.planpin.Value
            
            Range("CR" & contador).Value = FormularioDados.tipoestrutura.Value
            Range("CS" & contador).Value = CDbl(FormularioDados.modulo2.Value)
            Range("CT" & contador).Value = CDbl(FormularioDados.largmodulo2.Value)
            Range("CU" & contador).Value = FormularioDados.planpin.Value
            
            Range("DP" & contador).Value = FormularioDados.tipoestrutura.Value
            Range("DQ" & contador).Value = CDbl(FormularioDados.modulo3.Value)
            Range("DR" & contador).Value = CDbl(FormularioDados.largmodulo3.Value)
            Range("DS" & contador).Value = FormularioDados.planpin.Value
            
            Range("EN" & contador).Value = FormularioDados.tipoestrutura.Value
            Range("EO" & contador).Value = CDbl(FormularioDados.modulo4.Value)
            Range("EP" & contador).Value = CDbl(FormularioDados.largmodulo4.Value)
            Range("EQ" & contador).Value = FormularioDados.planpin.Value
        
            Range("FL" & contador).Value = FormularioDados.tipoestrutura.Value
            Range("FM" & contador).Value = CDbl(FormularioDados.modulo4.Value)
            Range("FN" & contador).Value = CDbl(FormularioDados.largmodulo4.Value)
            Range("FO" & contador).Value = FormularioDados.planpin.Value
            
            Range("GJ" & contador).Value = FormularioDados.tipoestrutura.Value
            Range("GK" & contador).Value = CDbl(FormularioDados.modulo4.Value)
            Range("GL" & contador).Value = CDbl(FormularioDados.largmodulo4.Value)
            Range("GM" & contador).Value = FormularioDados.planpin.Value
            
            Range("HH" & contador).Value = FormularioDados.tipoestrutura.Value
            Range("HI" & contador).Value = CDbl(FormularioDados.modulo4.Value)
            Range("HJ" & contador).Value = CDbl(FormularioDados.largmodulo4.Value)
            Range("HK" & contador).Value = FormularioDados.planpin.Value
        End If
        
        If FormularioDados.nrmodulos.Value = "8 Módulos" Then
            Range("BT" & contador).Value = FormularioDados.tipoestrutura.Value
            Range("BU" & contador).Value = CDbl(FormularioDados.modulo1.Value)
            Range("BV" & contador).Value = CDbl(FormularioDados.largmodulo1.Value)
            Range("BW" & contador).Value = FormularioDados.planpin.Value
            
            Range("CR" & contador).Value = FormularioDados.tipoestrutura.Value
            Range("CS" & contador).Value = CDbl(FormularioDados.modulo2.Value)
            Range("CT" & contador).Value = CDbl(FormularioDados.largmodulo2.Value)
            Range("CU" & contador).Value = FormularioDados.planpin.Value
            
            Range("DP" & contador).Value = FormularioDados.tipoestrutura.Value
            Range("DQ" & contador).Value = CDbl(FormularioDados.modulo3.Value)
            Range("DR" & contador).Value = CDbl(FormularioDados.largmodulo3.Value)
            Range("DS" & contador).Value = FormularioDados.planpin.Value
            
            Range("EN" & contador).Value = FormularioDados.tipoestrutura.Value
            Range("EO" & contador).Value = CDbl(FormularioDados.modulo4.Value)
            Range("EP" & contador).Value = CDbl(FormularioDados.largmodulo4.Value)
            Range("EQ" & contador).Value = FormularioDados.planpin.Value
        
            Range("FL" & contador).Value = FormularioDados.tipoestrutura.Value
            Range("FM" & contador).Value = CDbl(FormularioDados.modulo4.Value)
            Range("FN" & contador).Value = CDbl(FormularioDados.largmodulo4.Value)
            Range("FO" & contador).Value = FormularioDados.planpin.Value
            
            Range("GJ" & contador).Value = FormularioDados.tipoestrutura.Value
            Range("GK" & contador).Value = CDbl(FormularioDados.modulo4.Value)
            Range("GL" & contador).Value = CDbl(FormularioDados.largmodulo4.Value)
            Range("GM" & contador).Value = FormularioDados.planpin.Value
            
            Range("HH" & contador).Value = FormularioDados.tipoestrutura.Value
            Range("HI" & contador).Value = CDbl(FormularioDados.modulo4.Value)
            Range("HJ" & contador).Value = CDbl(FormularioDados.largmodulo4.Value)
            Range("HK" & contador).Value = FormularioDados.planpin.Value
            
            Range("IF" & contador).Value = FormularioDados.tipoestrutura.Value
            Range("IG" & contador).Value = CDbl(FormularioDados.modulo4.Value)
            Range("IH" & contador).Value = CDbl(FormularioDados.largmodulo4.Value)
            Range("II" & contador).Value = FormularioDados.planpin.Value
        End If
        
        Range("JD" & contador).Value = FormularioDados.complexidade.Value
        Range("KE" & contador).Value = FormularioDados.complexidade.Value
        
        If FormularioDados.tipomaq.Value = "Split" Then
            Range("JT" & contador).Value = "Split"
        Else
            Range("JT" & contador).Value = "Wall Mounted"
        End If
        
        Range("JS" & contador).Value = FormularioDados.qtdmaq.Value
        
        Range("JY" & contador).Value = FormularioDados.incendio.Value
        Range("KB" & contador).Value = FormularioDados.seguranca.Value
        
        Range("KJ" & contador).Value = CInt(FormularioDados.nrcolunas.Value)
        
        'RETIRADA QTD DE PAINEIS INTERLIGAÇÃO
        'Range("KK" & contador).Value = CInt(FormularioDados.paineisint.Value)
        
        If FormularioDados.whitemartins.Value = True Then
            Range("KZ" & contador).Value = "White Martins"
        Else
            Range("KZ" & contador).Value = "Outro"
        End If
        
        'Grava dados em variáveis Eng Mec
        HorLOM = Range("T" & contador).Value
        DurLOM = Range("U" & contador).Value
        HorLMM = Range("V" & contador).Value
        DurLMM = Range("W" & contador).Value
        HorPBS = Range("X" & contador).Value
        DurPBS = Range("Y" & contador).Value
        HorPPA = Range("Z" & contador).Value
        DurPPA = Range("AA" & contador).Value
        HorPCI = Range("AB" & contador).Value
        DurPCI = Range("AC" & contador).Value
        HorPCE = Range("AD" & contador).Value
        DurPCE = Range("AE" & contador).Value
        HorPAC = Range("AF" & contador).Value
        DurPAC = Range("AG" & contador).Value
        HorLCA = Range("AH" & contador).Value
        DurLCA = Range("AI" & contador).Value
        HorLAA = Range("AJ" & contador).Value
        DurLAA = Range("AK" & contador).Value
        HorLAM = Range("AL" & contador).Value
        DurLAM = Range("AM" & contador).Value
        HorLMA = Range("AN" & contador).Value
        DurLMA = Range("AO" & contador).Value
        HorPTR = Range("AP" & contador).Value
        DurPTR = Range("AQ" & contador).Value
        
        'Grava dados em variáveis Processos
        If FormularioDados.nrmodulos.Value = "1 Módulo" Then modulos = 1
        If FormularioDados.nrmodulos.Value = "2 Módulos" Then modulos = 2
        If FormularioDados.nrmodulos.Value = "3 Módulos" Then modulos = 3
        If FormularioDados.nrmodulos.Value = "4 Módulos" Then modulos = 4
        If FormularioDados.nrmodulos.Value = "5 Módulos" Then modulos = 5
        If FormularioDados.nrmodulos.Value = "6 Módulos" Then modulos = 6
        If FormularioDados.nrmodulos.Value = "7 Módulos" Then modulos = 7
        If FormularioDados.nrmodulos.Value = "8 Módulos" Then modulos = 8
        
        Hor531 = 3.1 * modulos
        If Hor531 > 7.6 Then
            Dur531 = 2
        Else
            Dur531 = 1
        End If
        If Hor531 > 15.2 Then Hor531 = 15.2
        
        Hor551 = 0.8 * modulos
        If Hor551 > 7.6 Then
            Dur551 = 2
        Else
            Dur551 = 1
        End If
        If Hor551 > 15.2 Then Hor551 = 15.2
        
        Hor561 = 0.8 * modulos
        If Hor561 > 7.6 Then
            Dur561 = 2
        Else
            Dur561 = 1
        End If
        If Hor561 > 15.2 Then Hor561 = 15.2

        Hor581 = 4 * modulos
        If Hor581 > 7.6 Then
            Dur581 = 2
        Else
            Dur581 = 1
        End If
        If Hor581 > 15.2 Then Hor581 = 15.2
        
        Hor585 = 4.8 * modulos
        If Hor585 > 7.6 Then
            Dur585 = 2
        Else
            Dur585 = 1
        End If
        If Hor585 > 15.2 Then Hor585 = 15.2
        
        Hor589 = 2 * modulos
        If Hor589 > 7.6 Then
            Dur589 = 2
        Else
            Dur589 = 1
        End If
        If Hor589 > 15.2 Then Hor589 = 15.2
        
        'Grava dados em variáveis Eng Elét
        HorPIL = Range("AR" & contador).Value
        DurPIL = Range("AS" & contador).Value
        HorPCL = Range("AT" & contador).Value
        DurPCL = Range("AU" & contador).Value
        HorCSM = Range("AV" & contador).Value
        DurCSM = Range("AW" & contador).Value
        HorLMC = Range("AX" & contador).Value
        DurLMC = Range("AY" & contador).Value
        HorPIN = Range("AZ" & contador).Value
        DurPIN = Range("BA" & contador).Value
        HorLMI = Range("BB" & contador).Value
        DurLMI = Range("BC" & contador).Value
        
        If Range("B" & contador).Value = "Sim" Then
            HorPSS = Range("BD" & contador).Value
            DurPSS = Range("BE" & contador).Value
            HorLMS = Range("BF" & contador).Value
            DurLMS = Range("BG" & contador).Value
        End If
        
        HorPBA = Range("BH" & contador).Value
        DurPBA = Range("BI" & contador).Value
        HorDIN = Range("BJ" & contador).Value
        DurDIN = Range("BK" & contador).Value
        HorLMT = Range("BL" & contador).Value
        DurLMT = Range("BM" & contador).Value
        HorLMB = Range("BN" & contador).Value
        DurLMB = Range("BO" & contador).Value
        HorLMD = Range("BP" & contador).Value
        DurLMD = Range("BQ" & contador).Value
        HorPRF = Range("BR" & contador).Value
        DurPRF = Range("BS" & contador).Value

        'Grava dados em variáveis Módulo 1
        HorCOR1 = Range("BX" & contador).Value
        DurCOR1 = Range("BY" & contador).Value
        HorFCH1 = Range("BZ" & contador).Value
        DurFCH1 = Range("CA" & contador).Value
        
        If FormularioDados.tipoestrutura.Value <> "Container Solar" And FormularioDados.tipoestrutura.Value <> "ESSW (mecânica)" Then
            HorPRB1 = Range("CB" & contador).Value
            DurPRB1 = Range("CC" & contador).Value
            HorSBA1 = Range("CD" & contador).Value
            DurSBA1 = Range("CE" & contador).Value
            HorPRE1 = Range("CF" & contador).Value
            DurPRE1 = Range("CG" & contador).Value
            HorSES1 = Range("CH" & contador).Value
            DurSES1 = Range("CI" & contador).Value
        End If
        
        HorEDF1 = Range("CJ" & contador).Value
        DurEDF1 = Range("CK" & contador).Value
        HorPIN1 = Range("CL" & contador).Value
        DurPIN1 = Range("CM" & contador).Value
        HorCHI1 = Range("CN" & contador).Value
        DurCHI1 = Range("CO" & contador).Value
        HorCHE1 = Range("CP" & contador).Value
        DurCHE1 = Range("CQ" & contador).Value
        
        'Grava dados em variáveis Módulo 2
        If FormularioDados.nrmodulos.Value <> "1 Módulo" Then
            HorCOR2 = Range("CV" & contador).Value
            DurCOR2 = Range("CW" & contador).Value
            HorFCH2 = Range("CX" & contador).Value
            DurFCH2 = Range("CY" & contador).Value
            HorPRB2 = Range("CZ" & contador).Value
            DurPRB2 = Range("DA" & contador).Value
            HorSBA2 = Range("DB" & contador).Value
            DurSBA2 = Range("DC" & contador).Value
            HorPRE2 = Range("DD" & contador).Value
            DurPRE2 = Range("DE" & contador).Value
            HorSES2 = Range("DF" & contador).Value
            DurSES2 = Range("DG" & contador).Value
            HorEDF2 = Range("DH" & contador).Value
            DurEDF2 = Range("DI" & contador).Value
            HorPIN2 = Range("DJ" & contador).Value
            DurPIN2 = Range("DK" & contador).Value
            HorCHI2 = Range("DL" & contador).Value
            DurCHI2 = Range("DM" & contador).Value
            HorCHE2 = Range("DN" & contador).Value
            DurCHE2 = Range("DO" & contador).Value
        End If
        
        'Grava dados em variáveis Módulo 3
        If FormularioDados.nrmodulos.Value <> "1 Módulo" Or FormularioDados.nrmodulos.Value <> "2 Módulos" Then
            HorCOR3 = Range("DT" & contador).Value
            DurCOR3 = Range("DU" & contador).Value
            HorFCH3 = Range("DV" & contador).Value
            DurFCH3 = Range("DW" & contador).Value
            HorPRB3 = Range("DX" & contador).Value
            DurPRB3 = Range("DY" & contador).Value
            HorSBA3 = Range("DZ" & contador).Value
            DurSBA3 = Range("EA" & contador).Value
            HorPRE3 = Range("EB" & contador).Value
            DurPRE3 = Range("EC" & contador).Value
            HorSES3 = Range("ED" & contador).Value
            DurSES3 = Range("EE" & contador).Value
            HorEDF3 = Range("EF" & contador).Value
            DurEDF3 = Range("EG" & contador).Value
            HorPIN3 = Range("EH" & contador).Value
            DurPIN3 = Range("EI" & contador).Value
            HorCHI3 = Range("EJ" & contador).Value
            DurCHI3 = Range("EK" & contador).Value
            HorCHE3 = Range("EL" & contador).Value
            DurCHE3 = Range("EM" & contador).Value
        End If
        
        'Grava dados em variáveis Módulo 4
        If FormularioDados.nrmodulos.Value <> "1 Módulo" Or FormularioDados.nrmodulos.Value <> "2 Módulos" Or FormularioDados.nrmodulos.Value <> "3 Módulos" Then
            HorCOR4 = Range("ER" & contador).Value
            DurCOR4 = Range("ES" & contador).Value
            HorFCH4 = Range("ET" & contador).Value
            DurFCH4 = Range("EU" & contador).Value
            HorPRB4 = Range("EV" & contador).Value
            DurPRB4 = Range("EW" & contador).Value
            HorSBA4 = Range("EX" & contador).Value
            DurSBA4 = Range("EY" & contador).Value
            HorPRE4 = Range("EZ" & contador).Value
            DurPRE4 = Range("FA" & contador).Value
            HorSES4 = Range("FB" & contador).Value
            DurSES4 = Range("FC" & contador).Value
            HorEDF4 = Range("FD" & contador).Value
            DurEDF4 = Range("FE" & contador).Value
            HorPIN4 = Range("FF" & contador).Value
            DurPIN4 = Range("FG" & contador).Value
            HorCHI4 = Range("FH" & contador).Value
            DurCHI4 = Range("FI" & contador).Value
            HorCHE4 = Range("FJ" & contador).Value
            DurCHE4 = Range("FK" & contador).Value
        End If

        'Grava dados em variáveis Módulo 5
        If FormularioDados.nrmodulos.Value <> "1 Módulo" Or FormularioDados.nrmodulos.Value <> "2 Módulos" Or FormularioDados.nrmodulos.Value <> "3 Módulos" Or FormularioDados.nrmodulos.Value <> "4 Módulos" Then
            HorCOR5 = Range("FP" & contador).Value
            DurCOR5 = Range("FQ" & contador).Value
            HorFCH5 = Range("FR" & contador).Value
            DurFCH5 = Range("FS" & contador).Value
            HorPRB5 = Range("FT" & contador).Value
            DurPRB5 = Range("FU" & contador).Value
            HorSBA5 = Range("FV" & contador).Value
            DurSBA5 = Range("FW" & contador).Value
            HorPRE5 = Range("FX" & contador).Value
            DurPRE5 = Range("FY" & contador).Value
            HorSES5 = Range("FZ" & contador).Value
            DurSES5 = Range("GA" & contador).Value
            HorEDF5 = Range("GB" & contador).Value
            DurEDF5 = Range("GC" & contador).Value
            HorPIN5 = Range("GD" & contador).Value
            DurPIN5 = Range("GE" & contador).Value
            HorCHI5 = Range("GF" & contador).Value
            DurCHI5 = Range("GG" & contador).Value
            HorCHE5 = Range("GH" & contador).Value
            DurCHE5 = Range("GI" & contador).Value
        End If
        
        'Grava dados em variáveis Módulo 6
        If FormularioDados.nrmodulos.Value <> "1 Módulo" Or FormularioDados.nrmodulos.Value <> "2 Módulos" Or FormularioDados.nrmodulos.Value <> "3 Módulos" Or FormularioDados.nrmodulos.Value <> "4 Módulos" Or FormularioDados.nrmodulos.Value <> "5 Módulos" Then
            HorCOR6 = Range("GN" & contador).Value
            DurCOR6 = Range("GO" & contador).Value
            HorFCH6 = Range("GP" & contador).Value
            DurFCH6 = Range("GQ" & contador).Value
            HorPRB6 = Range("GR" & contador).Value
            DurPRB6 = Range("GS" & contador).Value
            HorSBA6 = Range("GT" & contador).Value
            DurSBA6 = Range("GU" & contador).Value
            HorPRE6 = Range("GV" & contador).Value
            DurPRE6 = Range("GW" & contador).Value
            HorSES6 = Range("GX" & contador).Value
            DurSES6 = Range("GY" & contador).Value
            HorEDF6 = Range("GZ" & contador).Value
            DurEDF6 = Range("HA" & contador).Value
            HorPIN6 = Range("HB" & contador).Value
            DurPIN6 = Range("HC" & contador).Value
            HorCHI6 = Range("HD" & contador).Value
            DurCHI6 = Range("HE" & contador).Value
            HorCHE6 = Range("HF" & contador).Value
            DurCHE6 = Range("HG" & contador).Value
        End If
        
        'Grava dados em variáveis Módulo 7
        If FormularioDados.nrmodulos.Value <> "1 Módulo" Or FormularioDados.nrmodulos.Value <> "2 Módulos" Or FormularioDados.nrmodulos.Value <> "3 Módulos" Or FormularioDados.nrmodulos.Value <> "4 Módulos" Or FormularioDados.nrmodulos.Value <> "5 Módulos" Or FormularioDados.nrmodulos.Value <> "6 Módulos" Then
            HorCOR7 = Range("HL" & contador).Value
            DurCOR7 = Range("HM" & contador).Value
            HorFCH7 = Range("HN" & contador).Value
            DurFCH7 = Range("HO" & contador).Value
            HorPRB7 = Range("HP" & contador).Value
            DurPRB7 = Range("HQ" & contador).Value
            HorSBA7 = Range("HR" & contador).Value
            DurSBA7 = Range("HS" & contador).Value
            HorPRE7 = Range("HT" & contador).Value
            DurPRE7 = Range("HU" & contador).Value
            HorSES7 = Range("HV" & contador).Value
            DurSES7 = Range("HW" & contador).Value
            HorEDF7 = Range("HX" & contador).Value
            DurEDF7 = Range("HY" & contador).Value
            HorPIN7 = Range("HZ" & contador).Value
            DurPIN7 = Range("IA" & contador).Value
            HorCHI7 = Range("IB" & contador).Value
            DurCHI7 = Range("IC" & contador).Value
            HorCHE7 = Range("ID" & contador).Value
            DurCHE7 = Range("IE" & contador).Value
        End If
        
        'Grava dados em variáveis Módulo 8
        If FormularioDados.nrmodulos.Value <> "1 Módulo" Or FormularioDados.nrmodulos.Value <> "2 Módulos" Or FormularioDados.nrmodulos.Value <> "3 Módulos" Or FormularioDados.nrmodulos.Value <> "4 Módulos" Or FormularioDados.nrmodulos.Value <> "5 Módulos" Or FormularioDados.nrmodulos.Value <> "6 Módulos" Or FormularioDados.nrmodulos.Value <> "7 Módulos" Then
            HorCOR8 = Range("IJ" & contador).Value
            DurCOR8 = Range("IK" & contador).Value
            HorFCH8 = Range("IL" & contador).Value
            DurFCH8 = Range("IM" & contador).Value
            HorPRB8 = Range("IN" & contador).Value
            DurPRB8 = Range("IO" & contador).Value
            HorSBA8 = Range("IP" & contador).Value
            DurSBA8 = Range("IQ" & contador).Value
            HorPRE8 = Range("IR" & contador).Value
            DurPRE8 = Range("IS" & contador).Value
            HorSES8 = Range("IT" & contador).Value
            DurSES8 = Range("IU" & contador).Value
            HorEDF8 = Range("IV" & contador).Value
            DurEDF8 = Range("IW" & contador).Value
            HorPIN8 = Range("IX" & contador).Value
            DurPIN8 = Range("IY" & contador).Value
            HorCHI8 = Range("IZ" & contador).Value
            DurCHI8 = Range("JA" & contador).Value
            HorCHE8 = Range("JB" & contador).Value
            DurCHE8 = Range("JC" & contador).Value
        End If
        
        'Grava dados em variáveis Acessórios
        HorFAC = Range("JE" & contador).Value + Range("JG" & contador).Value
        DurFAC = Range("JF" & contador).Value
        HorFCA = Range("JH" & contador).Value
        DurFCA = Range("JI" & contador).Value
        HorMAM = Range("JJ" & contador).Value
        DurMAM = Range("LK" & contador).Value
        HorMAA = Range("JL" & contador).Value
        DurMAA = Range("JM" & contador).Value
        HorPRM = Range("JO" & contador).Value
        DurPRM = Range("JP" & contador).Value
        HorIST = Range("JQ" & contador).Value
        DurIST = Range("JR" & contador).Value
        
        HorMCL = Range("JU" & contador).Value
        DurMCL = Range("JV" & contador).Value
        
        If Range("E" & contador).Value = "Sim" Then
            HorMCM = Range("GE" & contador).Value
            DurMCM = Range("GF" & contador).Value
        End If
        
        HorMIN = Range("JZ" & contador).Value
        DurMIN = Range("KA" & contador).Value
        
        If FormularioDados.seguranca.Value <> "Não possui" Then
            HorMSS = Range("KC" & contador).Value
            DurMSS = Range("KD" & contador).Value
        End If
        
        'Grava dados em variáveis Elétromecânica
        HorFEQ = Range("KF" & contador).Value
        DurFEQ = Range("KG" & contador).Value
        HorLBA = Range("KH" & contador).Value
        DurLBA = Range("KI" & contador).Value
        HorINT = Range("KL" & contador).Value
        DurINT = Range("KM" & contador).Value
        HorTES = Range("KN" & contador).Value
        DurTES = Range("KO" & contador).Value
        HorINS = Range("KP" & contador).Value
        DurINS = Range("KQ" & contador).Value
        HorPEE = Range("KR" & contador).Value
        DurPEE = Range("KS" & contador).Value
        HorPEM = Range("KT" & contador).Value
        DurPEM = Range("KU" & contador).Value
        HorFEC = Range("KV" & contador).Value
        DurFEC = Range("KW" & contador).Value
        HorFEA = Range("KX" & contador).Value
        DurFEA = Range("KY" & contador).Value
        Exit For
    End If
Next
ActiveSheet.Protect

Call insere_tempos

End Sub
```

#### `insere_tempos` (Sub)

```vb
Sub insere_tempos()

Dim lastrow As Long, i As Long, j As Long

'Insere dados da tabela no resultado
resultado.Activate

If FormularioDados.tipoestrutura.Value = "Container Solar" Or FormularioDados.tipoestrutura.Value = "ESSW (mecânica)" Then
    Range("E4").Value = HorLOM
    Range("C4").Value = DurLOM
    Range("E8").Value = HorLMM
    Range("C8").Value = DurLMM
    Range("E10").Value = HorPBS
    Range("C10").Value = DurPBS
    Range("E13").Value = HorPPA
    Range("C13").Value = DurPPA
    Range("E16").Value = HorPCI
    Range("C16").Value = DurPCI
    Range("E19").Value = HorPCE
    Range("C19").Value = DurPCE
    Range("E22").Value = HorPAC
    Range("C22").Value = DurPAC
    Range("E24").Value = HorLCA
    Range("C24").Value = DurLCA
    Range("E27").Value = HorLAA
    Range("C27").Value = DurLAA
    Range("E30").Value = HorLAM
    Range("C30").Value = DurLAM
    Range("E33").Value = HorLMA
    Range("C33").Value = DurLMA
    Range("E37").Value = HorPTR
    Range("C37").Value = DurPTR
    Range("E41").Value = HorPIL
    Range("C41").Value = DurPIL
    Range("E45").Value = HorPCL
    Range("C45").Value = DurPCL
    Range("E47").Value = HorLMC
    Range("C47").Value = DurLMC
    Range("E49").Value = HorPIN
    Range("C49").Value = DurPIN
    Range("E51").Value = HorLMI
    Range("C51").Value = DurLMI
    
    If FormularioDados.seguranca.Value <> "Não possui" Then
        Range("E53").Value = HorPSS
        Range("C53").Value = DurPSS
        Range("E55").Value = HorLMS
        Range("C55").Value = DurLMS
    End If
    
    Range("E43").Value = HorLMT
    Range("C43").Value = DurLMT
    
    Range("E57").Value = HorDIN
    Range("C57").Value = DurDIN
    Range("E59").Value = HorLMD
    Range("C59").Value = DurLMD
    Range("E61").Value = HorPBA
    Range("C61").Value = DurPBA
    Range("E63").Value = HorLMB
    Range("C63").Value = DurLMB
    
    Range("E65").Value = HorPRF
    Range("C65").Value = DurPRF
    
    'Insere tempos Processos
    Range("E11").Value = Hor531
    Range("C11").Value = Dur531
    Range("E17").Value = Hor551
    Range("C17").Value = Dur551
    Range("E20").Value = Hor561
    Range("C20").Value = Dur561
    Range("E28").Value = Hor581
    Range("C28").Value = Dur581
    Range("E31").Value = Hor585
    Range("C31").Value = Dur585
    Range("E34").Value = Hor589
    Range("C34").Value = Dur589
    
    'Insere os tempos do Módulo 1
    Range("E69").Value = HorCOR1
    Range("C69").Value = DurCOR1
    
    'If FormularioDados.externo.Value = True Then
    'Alterar tempos da ESU - 705
    'Range("E66").Value = "0,1"
    'Range("C66").Value = 10
    'Range("A66").Value = 705
    'Range("B66").Value = "ESU - Envio para Subcontratação item sub"
    'Alterar tempos da ESU - 720
    'End If
        
    Range("E70").Value = HorFCH1
    Range("C70").Value = DurFCH1
    Range("E72").Value = HorEDF1
    Range("C72").Value = DurEDF1
    Range("E75").Value = HorPIN1
    Range("C75").Value = DurPIN1
    Range("E78").Value = HorCHI1
    Range("C78").Value = DurCHI1
    
    Range("E71").Value = HorCHE1
    Range("C71").Value = DurCHE1
    
    'Insere os tempos de Acessórios
    Range("E120").Value = HorFAC
    Range("C120").Value = DurFAC
    Range("E123").Value = HorFCA
    Range("C123").Value = DurFCA
    Range("E127").Value = HorMAM
    Range("C127").Value = DurMAM
    Range("E129").Value = HorMAA
    Range("C129").Value = DurMAA
    
    Range("E132").Value = HorPRM
    Range("C132").Value = DurPRM
    
    Range("E133").Value = HorIST
    Range("C133").Value = DurIST
    Range("E135").Value = HorMCL
    Range("C135").Value = DurMCL
    
    If FormularioDados.tipomaq.Value = "Roof Top" Then
        Range("E136").Value = HorMCM
        Range("C136").Value = DurMCM
    End If
    
    Range("E133").Value = Range("E133").Value + HorMIN
    Range("C133").Value = Range("C133").Value + DurMIN
    
    Range("E140").Value = HorFEQ
    Range("C140").Value = DurFEQ
    Range("E142").Value = HorLBA
    Range("C142").Value = DurLBA
    
    If FormularioDados.seguranca.Value <> "Não possui" Then
        Range("E142").Value = Range("E142").Value + HorMSS
        Range("C142").Value = Range("C142").Value + DurMSS
    End If
    
    Range("E145").Value = HorINT
    Range("C145").Value = DurINT
    Range("E147").Value = HorTES
    Range("C147").Value = DurTES
    Range("E149").Value = HorINS
    Range("C149").Value = DurINS
    Range("E150").Value = HorPEE
    Range("C150").Value = DurPEE
    Range("E151").Value = HorPEM
    Range("C151").Value = DurPEM
    Range("E152").Value = HorFEC
    Range("C152").Value = DurFEC
    Range("E130").Value = HorFEA
    Range("C130").Value = DurFEA
    
    'Arredondar horas
    Range("A4").Select
    Selection.End(xlDown).Select
    nr_linhas = Selection.Row
    For x = 4 To nr_linhas
        Range("C" & x).Value = Round(Range("C" & x).Value, 1)
        Range("E" & x).Value = Round(Range("E" & x).Value, 1)
    Next
    
    
    '###Elimina tarefas desnecessárias###
    
    'Sem FAT
    If FormularioDados.chkFilho.Value = True Then
        Rows("154:154").Select
        Selection.Delete Shift:=xlUp
    End If
    
    'Sem elétrica (ESSW mecânica)
    If FormularioDados.tipoestrutura.Value = "ESSW (mecânica)" Then
        Rows("131:152").Select
        Selection.Delete Shift:=xlUp
    End If
    
    'Sem TFS
    If FormularioDados.testesw.Value = False And FormularioDados.tipoestrutura.Value <> "ESSW (mecânica)" Then
        Rows("148:148").Select
        Selection.Delete Shift:=xlUp
    End If
    
    'Sem impressão identificações (ITA)
    If FormularioDados.Betim1310.Value = False Then
        Rows("143:143").Select
        Selection.Delete Shift:=xlUp
    End If
    
    'Sem Sistema segurança
    If FormularioDados.tipoestrutura.Value <> "ESSW (mecânica)" Then
        Rows("139:139").Select
        Selection.Delete Shift:=xlUp
    End If
    
    'Sem Montagem Incêndio (unificado com Instalações)
    If FormularioDados.tipoestrutura.Value <> "ESSW (mecânica)" Then
        Rows("137:138").Select
        Selection.Delete Shift:=xlUp
    End If
    
    'Sem casa de máquinas
    If FormularioDados.tipomaq.Value <> "Roof Top" And FormularioDados.tipoestrutura.Value <> "ESSW (mecânica)" Then
        Rows("136:136").Select
        Selection.Delete Shift:=xlUp
    End If
    
    'Sem climatização
    'If FormularioDados.tipomaq.Value = "Não possui" And FormularioDados.tipoestrutura.Value <> "ESSW (mecânica)" Then
    '    Rows("134:135").Select
    '    Selection.Delete Shift:=xlUp
    'End If
    
    'Sem FAC/FCA
'    If FormularioDados.Betim1310.Value = True Then
'        Rows("123:123").Select
'        Selection.Delete Shift:=xlUp
'        Rows("120:120").Select
'        Selection.Delete Shift:=xlUp
'    End If
    
    'Sem módulo 4
    If FormularioDados.nrmodulos.Value = "1 Módulo" Or FormularioDados.nrmodulos.Value = "2 Módulos" Or FormularioDados.nrmodulos.Value = "3 Módulos" Then
        Rows("107:118").Select
        Selection.Delete Shift:=xlUp
    End If
    
    'Sem módulo 3
    If FormularioDados.nrmodulos.Value = "1 Módulo" Or FormularioDados.nrmodulos.Value = "2 Módulos" Then
        Rows("95:106").Select
        Selection.Delete Shift:=xlUp
    End If
    
    'Sem módulo 2
    If FormularioDados.nrmodulos.Value = "1 Módulo" Then
        Rows("83:94").Select
        Selection.Delete Shift:=xlUp
    End If

    'Container Solar (sem chapeamento externo)
    Rows("79:81").Select
    Selection.Delete Shift:=xlUp
     
    'Sem SEI
    'If FormularioDados.tipoestrutura.Value = "Container Solar" And FormularioDados.Betim1310.Value = True Then
    '    Rows("77:77").Select
    '    Selection.Delete Shift:=xlUp
    'End If
    
    'Pintura externa
    'If FormularioDados.tipoestrutura.Value = "Container Solar" And FormularioDados.Betim1310.Value = True Then
    '    Rows("75:75").Select
    '    Selection.Delete Shift:=xlUp
    'End If
    
    'Sem Edificação
    'If FormularioDados.tipoestrutura.Value = "Container Solar" Then
    '     Rows("72:72").Select
    '     Selection.Delete Shift:=xlUp
    'End If
    
    'Sem Chaparia
    'If FormularioDados.tipoestrutura.Value = "Container Solar" And FormularioDados.Betim1310.Value = True Then
    '     Rows("70:70").Select
    '     Selection.Delete Shift:=xlUp
    'End If
    
    'Sem Corte e Solda
    'If FormularioDados.tipoestrutura.Value = "Container Solar" And FormularioDados.externo.Value = True Then
    '     Rows("67:67").Select
    '     Selection.Delete Shift:=xlUp
    'End If
    
    'Sem sistema segurança (engenharia)
    If FormularioDados.seguranca.Value = "Não possui" Or FormularioDados.tipoestrutura.Value = "ESSW (mecânica)" Then
        Rows("53:56").Select
        Selection.Delete Shift:=xlUp
    End If
    
    'Sem tarefa 577
    Rows("25:25").Select
    Selection.Delete Shift:=xlUp
    
    'Container Solar (engenharia)
    Rows("19:21").Select
    Selection.Delete Shift:=xlUp
    Rows("13:15").Select
    Selection.Delete Shift:=xlUp

Else
    'Insere dados da tabela no resultado quando for Módulos
    Range("E4").Value = HorLOM
    Range("C4").Value = DurLOM
    Range("E8").Value = HorLMM
    Range("C8").Value = DurLMM
    Range("E10").Value = HorPBS + HorPPA
    Range("C10").Value = DurPBS + DurPPA
    Range("E13").Value = HorPCI
    Range("C13").Value = DurPCI
    Range("E16").Value = HorPCE
    Range("C16").Value = DurPCE
    Range("E19").Value = HorPAC
    Range("C19").Value = DurPAC
    Range("E21").Value = HorLCA
    Range("C21").Value = DurLCA
    Range("E24").Value = HorLAA
    Range("C24").Value = DurLAA
    Range("E27").Value = HorLAM
    Range("C27").Value = DurLAM
    Range("E30").Value = HorLMA
    Range("C30").Value = DurLMA
    Range("E34").Value = HorPTR
    Range("C34").Value = DurPTR
    
    Range("E38").Value = HorPIL
    Range("C38").Value = DurPIL
    Range("E42").Value = HorPCL
    Range("C42").Value = DurPCL
    Range("E44").Value = HorLMC
    Range("C44").Value = DurLMC
    Range("E46").Value = HorPIN
    Range("C46").Value = DurPIN
    Range("E48").Value = HorLMI
    Range("C48").Value = DurLMI
    
    If FormularioDados.seguranca.Value <> "Não possui" Then
        Range("E50").Value = HorPSS
        Range("C50").Value = DurPSS
        Range("E52").Value = HorLMS
        Range("C52").Value = DurLMS
    End If
    
    Range("E54").Value = HorDIN
    Range("C54").Value = DurDIN
    Range("E56").Value = HorLMD
    Range("C56").Value = DurLMD
    Range("E58").Value = HorPBA
    Range("C58").Value = DurPBA
    Range("E60").Value = HorLMB
    Range("C60").Value = DurLMB
    
    Range("E40").Value = HorLMT
    Range("C40").Value = DurLMT
    Range("E62").Value = HorPRF
    'Range("C60").Value = DurPRF
    
     'Insere tempos Processos
    Range("E11").Value = Hor531
    Range("C11").Value = Dur531
    Range("E14").Value = Hor551
    Range("C14").Value = Dur551
    Range("E17").Value = Hor561
    Range("C17").Value = Dur561
    Range("E25").Value = Hor581
    Range("C25").Value = Dur581
    Range("E28").Value = Hor585
    Range("C28").Value = Dur585
    Range("E31").Value = Hor589
    Range("C31").Value = Dur589
    
    'Insere os tempos do Módulo 1
    Range("E66").Value = HorCOR1
    Range("C66").Value = DurCOR1
    Range("E68").Value = HorFCH1
    Range("C68").Value = DurFCH1
    Range("E69").Value = HorPRB1
    Range("C69").Value = DurPRB1
    Range("E70").Value = HorSBA1
    Range("C70").Value = DurSBA1
    Range("E71").Value = HorPRE1
    Range("C71").Value = DurPRE1
    Range("E72").Value = HorSES1
    Range("C72").Value = DurSES1
    Range("E73").Value = HorEDF1
    Range("C73").Value = DurEDF1
    Range("E75").Value = HorPIN1
    Range("C75").Value = DurPIN1
    Range("E78").Value = HorCHI1
    Range("C78").Value = DurCHI1
    Range("E81").Value = HorCHE1
    Range("C81").Value = DurCHE1
    
    'Insere os tempos do Módulo 2
    If FormularioDados.nrmodulos.Value <> "1 Módulo" Then
        Range("E84").Value = HorCOR2
        Range("C84").Value = DurCOR2
        Range("E86").Value = HorFCH2
        Range("C86").Value = DurFCH2
        Range("E87").Value = HorPRB2
        Range("C87").Value = DurPRB2
        Range("E88").Value = HorSBA2
        Range("C88").Value = DurSBA2
        Range("E89").Value = HorPRE2
        Range("C89").Value = DurPRE2
        Range("E90").Value = HorSES2
        Range("C90").Value = DurSES2
        Range("E91").Value = HorEDF2
        Range("C91").Value = DurEDF2
        Range("E93").Value = HorPIN2
        Range("C93").Value = DurPIN2
        Range("E96").Value = HorCHI2
        Range("C96").Value = DurCHI2
        Range("E99").Value = HorCHE2
        Range("C99").Value = DurCHE2
    End If
    
    'Insere os tempos do Módulo 3
    If FormularioDados.nrmodulos.Value <> "1 Módulo" Or FormularioDados.nrmodulos.Value <> "2 Módulos" Then
        Range("E102").Value = HorCOR3
        Range("C102").Value = DurCOR3
        Range("E104").Value = HorFCH3
        Range("C104").Value = DurFCH3
        Range("E105").Value = HorPRB3
        Range("C105").Value = DurPRB3
        Range("E106").Value = HorSBA3
        Range("C106").Value = DurSBA3
        Range("E107").Value = HorPRE3
        Range("C107").Value = DurPRE3
        Range("E108").Value = HorSES3
        Range("C108").Value = DurSES3
        Range("E109").Value = HorEDF3
        Range("C109").Value = DurEDF3
        Range("E111").Value = HorPIN3
        Range("C111").Value = DurPIN3
        Range("E114").Value = HorCHI3
        Range("C114").Value = DurCHI3
        Range("E117").Value = HorCHE3
        Range("C117").Value = DurCHE3
    End If
    
    'Insere os tempos do Módulo 4
    If FormularioDados.nrmodulos.Value <> "1 Módulo" Or FormularioDados.nrmodulos.Value <> "2 Módulos" Or FormularioDados.nrmodulos.Value <> "3 Módulos" Then
        Range("E120").Value = HorCOR4
        Range("C120").Value = DurCOR4
        Range("E122").Value = HorFCH4
        Range("C122").Value = DurFCH4
        Range("E123").Value = HorPRB4
        Range("C123").Value = DurPRB4
        Range("E124").Value = HorSBA4
        Range("C124").Value = DurSBA4
        Range("E125").Value = HorPRE4
        Range("C125").Value = DurPRE4
        Range("E126").Value = HorSES4
        Range("C126").Value = DurSES4
        Range("E127").Value = HorEDF4
        Range("C127").Value = DurEDF4
        Range("E129").Value = HorPIN4
        Range("C129").Value = DurPIN4
        Range("E132").Value = HorCHI4
        Range("C132").Value = DurCHI4
        Range("E135").Value = HorCHE4
        Range("C135").Value = DurCHE4
    End If
    
    'Insere os tempos do Módulo 5
    If FormularioDados.nrmodulos.Value <> "1 Módulo" Or FormularioDados.nrmodulos.Value <> "2 Módulos" Or FormularioDados.nrmodulos.Value <> "3 Módulos" Or FormularioDados.nrmodulos.Value <> "4 Módulos" Then
        Range("E138").Value = HorCOR5
        Range("C138").Value = DurCOR5
        Range("E140").Value = HorFCH5
        Range("C140").Value = DurFCH5
        Range("E141").Value = HorPRB5
        Range("C141").Value = DurPRB5
        Range("E142").Value = HorSBA5
        Range("C142").Value = DurSBA5
        Range("E143").Value = HorPRE5
        Range("C143").Value = DurPRE5
        Range("E144").Value = HorSES5
        Range("C144").Value = DurSES5
        Range("E145").Value = HorEDF5
        Range("C145").Value = DurEDF5
        Range("E147").Value = HorPIN5
        Range("C147").Value = DurPIN5
        Range("E150").Value = HorCHI5
        Range("C150").Value = DurCHI5
        Range("E153").Value = HorCHE5
        Range("C153").Value = DurCHE5
    End If
    
    'Insere os tempos do Módulo 6
    If FormularioDados.nrmodulos.Value <> "1 Módulo" Or FormularioDados.nrmodulos.Value <> "2 Módulos" Or FormularioDados.nrmodulos.Value <> "3 Módulos" Or FormularioDados.nrmodulos.Value <> "4 Módulos" Or FormularioDados.nrmodulos.Value <> "5 Módulos" Then
        Range("E156").Value = HorCOR6
        Range("C156").Value = DurCOR6
        Range("E158").Value = HorFCH6
        Range("C158").Value = DurFCH6
        Range("E159").Value = HorPRB6
        Range("C159").Value = DurPRB6
        Range("E160").Value = HorSBA6
        Range("C160").Value = DurSBA6
        Range("E161").Value = HorPRE6
        Range("C161").Value = DurPRE6
        Range("E162").Value = HorSES6
        Range("C162").Value = DurSES6
        Range("E163").Value = HorEDF6
        Range("C163").Value = DurEDF6
        Range("E165").Value = HorPIN6
        Range("C165").Value = DurPIN6
        Range("E168").Value = HorCHI6
        Range("C168").Value = DurCHI6
        Range("E171").Value = HorCHE6
        Range("C171").Value = DurCHE6
    End If
    
    'Insere os tempos do Módulo 7
    If FormularioDados.nrmodulos.Value <> "1 Módulo" Or FormularioDados.nrmodulos.Value <> "2 Módulos" Or FormularioDados.nrmodulos.Value <> "3 Módulos" Or FormularioDados.nrmodulos.Value <> "4 Módulos" Or FormularioDados.nrmodulos.Value <> "5 Módulos" Or FormularioDados.nrmodulos.Value <> "6 Módulos" Then
        Range("E174").Value = HorCOR7
        Range("C174").Value = DurCOR7
        Range("E176").Value = HorFCH7
        Range("C176").Value = DurFCH7
        Range("E177").Value = HorPRB7
        Range("C177").Value = DurPRB7
        Range("E178").Value = HorSBA7
        Range("C178").Value = DurSBA7
        Range("E179").Value = HorPRE7
        Range("C179").Value = DurPRE7
        Range("E180").Value = HorSES7
        Range("C180").Value = DurSES7
        Range("E181").Value = HorEDF7
        Range("C181").Value = DurEDF7
        Range("E183").Value = HorPIN7
        Range("C183").Value = DurPIN7
        Range("E186").Value = HorCHI7
        Range("C186").Value = DurCHI7
        Range("E189").Value = HorCHE7
        Range("C189").Value = DurCHE7
    End If
    
    'Insere os tempos do Módulo 8
    If FormularioDados.nrmodulos.Value <> "1 Módulo" Or FormularioDados.nrmodulos.Value <> "2 Módulos" Or FormularioDados.nrmodulos.Value <> "3 Módulos" Or FormularioDados.nrmodulos.Value <> "4 Módulos" Or FormularioDados.nrmodulos.Value <> "5 Módulos" Or FormularioDados.nrmodulos.Value <> "6 Módulos" Or FormularioDados.nrmodulos.Value <> "7 Módulos" Then
        Range("E192").Value = HorCOR8
        Range("C192").Value = DurCOR8
        Range("E194").Value = HorFCH8
        Range("C194").Value = DurFCH8
        Range("E195").Value = HorPRB8
        Range("C195").Value = DurPRB8
        Range("E196").Value = HorSBA8
        Range("C196").Value = DurSBA8
        Range("E197").Value = HorPRE8
        Range("C197").Value = DurPRE8
        Range("E198").Value = HorSES8
        Range("C198").Value = DurSES8
        Range("E199").Value = HorEDF8
        Range("C199").Value = DurEDF8
        Range("E201").Value = HorPIN8
        Range("C201").Value = DurPIN8
        Range("E204").Value = HorCHI8
        Range("C204").Value = DurCHI8
        Range("E207").Value = HorCHE8
        Range("C207").Value = DurCHE8
    End If
    
    
    Range("E210").Value = HorFAC
    Range("C210").Value = DurFAC
    Range("E213").Value = HorFCA
    Range("C213").Value = DurFCA
    Range("E217").Value = HorMAM
    Range("C217").Value = DurMAM
    Range("E219").Value = HorMAA
    Range("C219").Value = DurMAA
    
    Range("E222").Value = HorPRM
    Range("C222").Value = DurPRM
    Range("E223").Value = HorIST
    Range("C223").Value = DurIST
    Range("E225").Value = HorMCL
    Range("C225").Value = DurMCL
    
    If FormularioDados.tipomaq.Value = "Roof Top" Then
        Range("E226").Value = HorMCM
        Range("C226").Value = DurMCM
    End If
    
    Range("E223").Value = Range("E223").Value + HorMIN
    Range("C223").Value = Range("C223").Value + DurMIN
    
    Range("E230").Value = HorFEQ
    Range("C230").Value = DurFEQ
    Range("E232").Value = HorLBA
    Range("C232").Value = DurLBA
    
    If FormularioDados.seguranca.Value <> "Não possui" Then
        Range("E232").Value = Range("E232").Value + HorMSS
        Range("C232").Value = Range("C232").Value + DurMSS
    End If
    
    Range("E235").Value = HorINT
    Range("C235").Value = DurINT
    Range("E237").Value = HorTES
    Range("C237").Value = DurTES
    Range("E239").Value = HorINS
    Range("C239").Value = DurINS
    Range("E240").Value = HorPEE
    Range("C240").Value = DurPEE
    Range("E241").Value = HorPEM
    Range("C241").Value = DurPEM
    Range("E242").Value = HorFEC
    Range("C242").Value = DurFEC
    Range("E220").Value = HorFEA
    Range("C220").Value = DurFEA
    
    'Arredondar horas
    Range("A4").Select
    Selection.End(xlDown).Select
    nr_linhas = Selection.Row
    For x = 4 To nr_linhas
        Range("C" & x).Value = Round(Range("C" & x).Value, 1)
        Range("E" & x).Value = Round(Range("E" & x).Value, 1)
    Next
    
    'Inverte EDF e PIN se for Betim
    If FormularioDados.Betim1310.Value = True Then
    
        'Módulo 1
        Range("A73:E74").Select
        Application.CutCopyMode = False
        Selection.Copy
        Range("A74").Select
        ActiveSheet.Paste
        Range("A73").Value = "750"
        Range("B73").Value = "PIN - Pintura Externa"
        Range("C73").Value = DurPIN1
        Range("E73").Value = Round(HorPIN1, 1)
        Range("A74").Value = "751"
        
        'Módulo 2
        If FormularioDados.nrmodulos.Value <> "1 Módulo" Then
            Range("A91:E92").Select
            Application.CutCopyMode = False
            Selection.Copy
            Range("A92").Select
            ActiveSheet.Paste
            Range("A91").Value = "750"
            Range("B91").Value = "PIN - Pintura Externa"
            Range("C91").Value = DurPIN2
            Range("E91").Value = Round(HorPIN2, 1)
            Range("A92").Value = "751"
        End If
        
        'Módulo 3
        If FormularioDados.nrmodulos.Value <> "1 Módulo" Or FormularioDados.nrmodulos.Value <> "2 Módulos" Then
            Range("A109:E110").Select
            Application.CutCopyMode = False
            Selection.Copy
            Range("A110").Select
            ActiveSheet.Paste
            Range("A109").Value = "750"
            Range("B109").Value = "PIN - Pintura Externa"
            Range("C109").Value = DurPIN3
            Range("E109").Value = Round(HorPIN3, 1)
            Range("A110").Value = "751"
        End If
        
        'Módulo 4
        If FormularioDados.nrmodulos.Value <> "1 Módulo" Or FormularioDados.nrmodulos.Value <> "2 Módulos" Or FormularioDados.nrmodulos.Value <> "3 Módulos" Then
            Range("A127:E128").Select
            Application.CutCopyMode = False
            Selection.Copy
            Range("A128").Select
            ActiveSheet.Paste
            Range("A127").Value = "750"
            Range("B127").Value = "PIN - Pintura Externa"
            Range("C127").Value = DurPIN4
            Range("E127").Value = Round(HorPIN4, 1)
            Range("A128").Value = "751"
        End If
        
        'Módulo 5
        If FormularioDados.nrmodulos.Value <> "1 Módulo" Or FormularioDados.nrmodulos.Value <> "2 Módulos" Or FormularioDados.nrmodulos.Value <> "3 Módulos" Or FormularioDados.nrmodulos.Value <> "4 Módulos" Then
            Range("A145:E146").Select
            Application.CutCopyMode = False
            Selection.Copy
            Range("A146").Select
            ActiveSheet.Paste
            Range("A145").Value = "750"
            Range("B145").Value = "PIN - Pintura Externa"
            Range("C145").Value = DurPIN5
            Range("E145").Value = Round(HorPIN5, 1)
            Range("A146").Value = "751"
        End If
        
        'Módulo 6
        If FormularioDados.nrmodulos.Value <> "1 Módulo" Or FormularioDados.nrmodulos.Value <> "2 Módulos" Or FormularioDados.nrmodulos.Value <> "3 Módulos" Or FormularioDados.nrmodulos.Value <> "4 Módulos" Or FormularioDados.nrmodulos.Value <> "5 Módulos" Then
            Range("A163:E164").Select
            Application.CutCopyMode = False
            Selection.Copy
            Range("A164").Select
            ActiveSheet.Paste
            Range("A163").Value = "750"
            Range("B163").Value = "PIN - Pintura Externa"
            Range("C163").Value = DurPIN6
            Range("E163").Value = Round(HorPIN6, 1)
            Range("A164").Value = "751"
        End If
        
        'Módulo 7
        If FormularioDados.nrmodulos.Value <> "1 Módulo" Or FormularioDados.nrmodulos.Value <> "2 Módulos" Or FormularioDados.nrmodulos.Value <> "3 Módulos" Or FormularioDados.nrmodulos.Value <> "4 Módulos" Or FormularioDados.nrmodulos.Value <> "5 Módulos" Or FormularioDados.nrmodulos.Value <> "6 Módulos" Then
            Range("A181:E182").Select
            Application.CutCopyMode = False
            Selection.Copy
            Range("A182").Select
            ActiveSheet.Paste
            Range("A181").Value = "750"
            Range("B181").Value = "PIN - Pintura Externa"
            Range("C181").Value = DurPIN7
            Range("E181").Value = Round(HorPIN7, 1)
            Range("A182").Value = "751"
        End If
        
        'Módulo 8
        If FormularioDados.nrmodulos.Value <> "1 Módulo" Or FormularioDados.nrmodulos.Value <> "2 Módulos" Or FormularioDados.nrmodulos.Value <> "3 Módulos" Or FormularioDados.nrmodulos.Value <> "4 Módulos" Or FormularioDados.nrmodulos.Value <> "5 Módulos" Or FormularioDados.nrmodulos.Value <> "6 Módulos" Or FormularioDados.nrmodulos.Value <> "7 Módulos" Then
            Range("A199:E200").Select
            Application.CutCopyMode = False
            Selection.Copy
            Range("A200").Select
            ActiveSheet.Paste
            Range("A199").Value = "750"
            Range("B199").Value = "PIN - Pintura Externa"
            Range("C199").Value = DurPIN8
            Range("E199").Value = Round(HorPIN8, 1)
            Range("A200").Value = "751"
        End If
    
    End If
    
    '###Elimina tarefas desnecessárias###
      
    'Sem FAT
    If FormularioDados.chkFilho.Value = True Then
        Rows("244:244").Select
        Selection.Delete Shift:=xlUp
    End If
    
    'Sem TFS
    If FormularioDados.testesw.Value = False Then
        Rows("238:238").Select
        Selection.Delete Shift:=xlUp
    End If
    
    'Sem impressão identificações (ITA)
    If FormularioDados.Betim1310.Value = False Then
        Rows("233:233").Select
        Selection.Delete Shift:=xlUp
    End If
    
    'Sem Sistema segurança (unificado com Leito/Bandej)
    Rows("229:229").Select
    Selection.Delete Shift:=xlUp
    
    'Sem Montagem Incêndio (unificado com Montagem Instalações)
    Rows("227:228").Select
    Selection.Delete Shift:=xlUp
    
    'Sem casa de máquinas
    If FormularioDados.tipomaq.Value <> "Roof Top" Then
        Rows("226:226").Select
        Selection.Delete Shift:=xlUp
    End If
    
    'Sem climatização
    'If FormularioDados.tipomaq.Value = "Não possui" Then
    '    Rows("224:225").Select
    '    Selection.Delete Shift:=xlUp
    'End If
    
    'Sem módulo 8
    If FormularioDados.nrmodulos.Value = "1 Módulo" Or FormularioDados.nrmodulos.Value = "2 Módulos" Or FormularioDados.nrmodulos.Value = "3 Módulos" Or FormularioDados.nrmodulos.Value = "4 Módulos" Or FormularioDados.nrmodulos.Value = "5 Módulos" Or FormularioDados.nrmodulos.Value = "6 Módulos" Or FormularioDados.nrmodulos.Value = "7 Módulos" Then
        Rows("191:208").Select
        Selection.Delete Shift:=xlUp
    Else
        Rows("208:208").Select
        Selection.Delete Shift:=xlUp
        'If FormularioDados.Betim1310.Value = False Then
            Rows("193:193").Select
            Selection.Delete Shift:=xlUp
        'End If
    End If
    
    'Sem módulo 7
    If FormularioDados.nrmodulos.Value = "1 Módulo" Or FormularioDados.nrmodulos.Value = "2 Módulos" Or FormularioDados.nrmodulos.Value = "3 Módulos" Or FormularioDados.nrmodulos.Value = "4 Módulos" Or FormularioDados.nrmodulos.Value = "5 Módulos" Or FormularioDados.nrmodulos.Value = "6 Módulos" Then
        Rows("173:190").Select
        Selection.Delete Shift:=xlUp
    Else
        Rows("190:190").Select
        Selection.Delete Shift:=xlUp
        'If FormularioDados.Betim1310.Value = False Then
            Rows("175:175").Select
            Selection.Delete Shift:=xlUp
        'End If
    End If
    
    'Sem módulo 6
    If FormularioDados.nrmodulos.Value = "1 Módulo" Or FormularioDados.nrmodulos.Value = "2 Módulos" Or FormularioDados.nrmodulos.Value = "3 Módulos" Or FormularioDados.nrmodulos.Value = "4 Módulos" Or FormularioDados.nrmodulos.Value = "5 Módulos" Then
        Rows("155:172").Select
        Selection.Delete Shift:=xlUp
    Else
        Rows("172:172").Select
        Selection.Delete Shift:=xlUp
        'If FormularioDados.Betim1310.Value = False Then
            Rows("157:157").Select
            Selection.Delete Shift:=xlUp
        'End If
    End If
    
    'Sem módulo 5
    If FormularioDados.nrmodulos.Value = "1 Módulo" Or FormularioDados.nrmodulos.Value = "2 Módulos" Or FormularioDados.nrmodulos.Value = "3 Módulos" Or FormularioDados.nrmodulos.Value = "4 Módulos" Then
        Rows("137:154").Select
        Selection.Delete Shift:=xlUp
    Else
        Rows("154:154").Select
        Selection.Delete Shift:=xlUp
        'If FormularioDados.Betim1310.Value = False Then
            Rows("139:139").Select
            Selection.Delete Shift:=xlUp
        'End If
    End If
        
    'Sem módulo 4
    If FormularioDados.nrmodulos.Value = "1 Módulo" Or FormularioDados.nrmodulos.Value = "2 Módulos" Or FormularioDados.nrmodulos.Value = "3 Módulos" Then
        Rows("119:136").Select
        Selection.Delete Shift:=xlUp
    Else
        Rows("136:136").Select
        Selection.Delete Shift:=xlUp
        'If FormularioDados.Betim1310.Value = False Then
            Rows("121:121").Select
            Selection.Delete Shift:=xlUp
        'End If
    End If
    
    'Sem módulo 3
    If FormularioDados.nrmodulos.Value = "1 Módulo" Or FormularioDados.nrmodulos.Value = "2 Módulos" Then
        Rows("101:118").Select
        Selection.Delete Shift:=xlUp
    Else
        Rows("118:118").Select
        Selection.Delete Shift:=xlUp
        'If FormularioDados.Betim1310.Value = False Then
            Rows("103:103").Select
            Selection.Delete Shift:=xlUp
        'End If
    End If
    
    'Sem módulo 2
    If FormularioDados.nrmodulos.Value = "1 Módulo" Then
        Rows("83:100").Select
        Selection.Delete Shift:=xlUp
    Else
        Rows("100:100").Select
        Selection.Delete Shift:=xlUp
        'If FormularioDados.Betim1310.Value = False Then
            Rows("85:85").Select
            Selection.Delete Shift:=xlUp
        'End If
    End If
    
    'Elimina REP
    Rows("82:82").Select
    Selection.Delete Shift:=xlUp
    
    'Sem PRE/SES (Móvel)
    If FormularioDados.tipoestrutura.Value = "Móvel" Then
        Rows("71:72").Select
        Selection.Delete Shift:=xlUp
    End If

    'Elimina a tarefa de ESU quando for 1 módulo
    If FormularioDados.nrmodulos.Value = "1 Módulo" Or FormularioDados.nrmodulos.Value = "2 Módulos" Or FormularioDados.nrmodulos.Value = "3 Módulos" Or FormularioDados.nrmodulos.Value = "4 Módulos" Or FormularioDados.nrmodulos.Value = "5 Módulos" Or FormularioDados.nrmodulos.Value = "6 Módulos" Or FormularioDados.nrmodulos.Value = "7 Módulos" Or FormularioDados.nrmodulos.Value = "8 Módulos" Then
        Rows("67:67").Select
        Selection.Delete Shift:=xlUp
    End If
    
    'Sem sistema segurança (engenharia)
    If FormularioDados.seguranca.Value = "Não possui" Then
        Rows("50:53").Select
        Selection.Delete Shift:=xlUp
    End If
    
    'Sem tarefa 577
    Rows("22:22").Select
    Selection.Delete Shift:=xlUp
    
End If

'Preenche nomenclatura
If FormularioDados.tipoestrutura.Value = "Container Solar" Then Range("F3").Value = "ESW"
If FormularioDados.tipoestrutura.Value = "ESSW (mecânica)" Then Range("F3").Value = "ESSW"
If FormularioDados.tipoestrutura.Value = "Móvel" Then Range("F3").Value = "ELW MÓVEL"
If FormularioDados.tipoestrutura.Value = "Semimóvel" Then Range("F3").Value = "ELW SEMIMÓVEL"
If FormularioDados.tipoestrutura.Value = "Modular" Then Range("F3").Value = "ELW MODULAR"
If FormularioDados.tipoestrutura.Value = "Fixo" Then Range("F3").Value = "ELW FIXO"
If FormularioDados.tipoestrutura.Value = "Embarcado" Then Range("F3").Value = "ELW EMBARCADO"

Dim conta As Integer
Dim linha_tarefa As Integer

For conta = 4 To 245
    If Range("A" & conta).Value = 885 Then
        Range("A" & conta).Select
        linha_tarefa = Selection.Row
        Exit For
    End If
Next

'Adicionar tarefa de TFS 1306
If FormularioDados.Betim1310.Value = False And FormularioDados.programacaoreles.Value = True Then
    'ADD tempos para tarefa de reles
        Sheets("Template").Select
        ActiveSheet.Unprotect
        Sheets("Resultado").Select
        ActiveSheet.Unprotect
        Sheets("Template").Select
        Range("A155:E155").Select
        Range("E155").Activate
        Selection.Copy
        Sheets("Resultado").Select
        Range("A" & linha_tarefa + 1).Select
        Selection.EntireRow.Insert
End If

If FormularioDados.tipoestrutura.Value = "ESSW (mecânica)" Then

    'Cria tabela de totais de horas
    Range("L3").Select
    ActiveCell.FormulaR1C1 = "ENG"
    Range("M3").Select
    Selection.FormulaArray = _
        "=SUM((R4C5:R160C5)*(R4C5:R160C5<>0.1)*(R4C1:R160C1<=702)*(R4C1:R160C1<>531)*(R4C1:R160C1<>551)*(R4C1:R160C1<>561)*(R4C1:R160C1<>581)*(R4C1:R160C1<>585)*(R4C1:R160C1<>589))"
    Selection.NumberFormat = "#,##0.0"
    Range("L4").Select
    ActiveCell.FormulaR1C1 = "MEC"
    Range("M4").Select
    Selection.FormulaArray = _
        "=SUM((R4C5:R160C5)*(R4C5:R160C5<>0.1)*(R4C1:R160C1<=798)*(R4C1:R160C1>=705)*(R4C1:R160C1<>754)*(R4C1:R160C1<>755)*(R4C1:R160C1<>765)*(R4C1:R160C1<>793)*(R4C1:R160C1<>794))+SUM((R4C5:R160C5)*(R4C5:R160C5<>0.1)*(R4C1:R160C1=894))"
    Selection.NumberFormat = "#,##0.0"
'    Range("L5").Select
'    ActiveCell.FormulaR1C1 = "ELE"
'    Range("M5").Select
'    Selection.FormulaArray = _
'        "=SUM((R4C5:R160C5)*(R4C5:R160C5<>0.1)*(R4C1:R160C1<=893)*(R4C1:R160C1>=799)*(R4C1:R160C1<>810)*(R4C1:R160C1<>828)*(R4C1:R160C1<>838)*(R4C1:R160C1<>858)*(R4C1:R160C1<>868))+SUM((R4C5:R160C5)*(R4C5:R160C5<>0.1)*(R4C1:R160C1=895))"
    Range("L5").Select
    ActiveCell.FormulaR1C1 = "TOTAL"
    Range("M5").Select
    Selection.FormulaArray = "=SUM(R[-3]C:R[-1]C)"
    Selection.NumberFormat = "#,##0.0"
    Range("L5:M5").Select
    Selection.Font.Bold = True
    Range("L3:M5").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
Else
    'Cria tabela de totais de horas
    Range("L3").Select
    ActiveCell.FormulaR1C1 = "ENG"
    Range("M3").Select
    Selection.FormulaArray = _
        "=SUM((R4C5:R255C5)*(R4C5:R255C5<>0.1)*(R4C1:R255C1<=702)*(R4C1:R255C1<>531)*(R4C1:R255C1<>551)*(R4C1:R255C1<>561)*(R4C1:R255C1<>581)*(R4C1:R255C1<>585)*(R4C1:R255C1<>589))"
    Selection.NumberFormat = "#,##0.0"
    Range("L4").Select
    ActiveCell.FormulaR1C1 = "MEC"
    Range("M4").Select
    Selection.FormulaArray = _
        "=SUM((R4C5:R255C5)*(R4C5:R255C5<>0.1)*(R4C1:R255C1<=798)*(R4C1:R255C1>=705)*(R4C1:R255C1<>754)*(R4C1:R255C1<>755)*(R4C1:R255C1<>765)*(R4C1:R255C1<>793)*(R4C1:R255C1<>794))+SUM((R4C5:R255C5)*(R4C5:R255C5<>0.1)*(R4C1:R255C1=894))"
    Selection.NumberFormat = "#,##0.0"
    Range("L5").Select
    ActiveCell.FormulaR1C1 = "ELE"
    Range("M5").Select
    Selection.FormulaArray = _
        "=SUM((R4C5:R255C5)*(R4C5:R255C5<>0.1)*(R4C1:R255C1<=893)*(R4C1:R255C1>=799)*(R4C1:R255C1<>810)*(R4C1:R255C1<>828)*(R4C1:R255C1<>838)*(R4C1:R255C1<>858)*(R4C1:R255C1<>868))+SUM((R4C5:R255C5)*(R4C5:R255C5<>0.1)*(R4C1:R255C1=895))"
    Selection.NumberFormat = "#,##0.0"
    Range("L6").Select
    ActiveCell.FormulaR1C1 = "TOTAL"
    Range("M6").Select
    Selection.FormulaArray = "=SUM(R[-3]C:R[-1]C)"
    Selection.NumberFormat = "#,##0.0"
    Range("L6:M6").Select
    Selection.Font.Bold = True
    Range("L3:M6").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
End If

'Seleciona resultado
Range("A4:E4").Select
Range(Selection, Selection.End(xlDown)).Select

'################################# BETIM ################################################

'If FormularioDados.tipoestrutura.Value = "Container Solar" And FormularioDados.Betim1310.Value = True And FormularioDados.programacaoreles.Value = True Then
'    ' Adicionar tarefas de se necessário reles
'        Sheets("Template").Select
'        ActiveSheet.Unprotect
'        Sheets("Resultado").Select
'        ActiveSheet.Unprotect
'        Sheets("Template").Select
'        Range("A155:E155").Select
'        Range("E155").Activate
'        Selection.Copy
'        Sheets("Resultado").Select
'        Range("A" & linha_tarefa + 1).Select
'        ActiveSheet.Unprotect
'        Selection.EntireRow.Insert
'End If

'################################# FIM BETIM ################################################

'Eliminar no 1306 tarefas de ENG se necessário

If FormularioDados.SemEng.Value = True And FormularioDados.tipoestrutura.Value = "Container Solar" Then
   ActiveSheet.Unprotect
   Range("A4:E57").Select
   Selection.Delete Shift:=xlUp
   ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
End If

If FormularioDados.SemEng.Value = True And FormularioDados.tipoestrutura.Value = "Móvel" Then
   ActiveSheet.Unprotect
   Range("A4:E60").Select
   Selection.Delete Shift:=xlUp
   ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
End If

If FormularioDados.proBTI.Value = True Then
        
    For conta = 4 To 255
        If Range("A" & conta).Value = 900 Or Range("A" & conta).Value = 899 Then
            Range("A" & conta).Select
            linha_tarefa = Selection.Row
            Exit For
        End If
    Next
    
    'Adicionar tarefa de EXP
    Sheets("Resultado").Select
    ActiveSheet.Unprotect
    Sheets("Template").Select
    ActiveSheet.Unprotect
    Range("A158:E158").Select
    Selection.Copy
    Sheets("Resultado").Select
    Range("A" & linha_tarefa + 1).Select
    Selection.EntireRow.Insert
End If


Sheets("Template").Select
ActiveSheet.Protect
Sheets("Resultado").Select
ActiveSheet.Protect

Application.ScreenUpdating = True

End Sub
```

#### `tempos_skid` (Sub)

```vb
Sub tempos_skid()

Application.ScreenUpdating = False

Dim HorPBSsKid As Double
Dim DurPBSskid As Integer
Dim HorBaciaCont As Integer
Dim HorPisoRem As Integer

If FormularioDados.txtPEP.Value <> "" Then
    ThisWorkbook.Application.StatusBar = "PEP " & FormularioDados.txtPEP.Value & ": Calculando tempos"
Else
    ThisWorkbook.Application.StatusBar = "Calculando tempos"
End If

ActiveSheet.Unprotect

'Limpa aba de resultados
Rows("4:245").Select
Selection.Delete Shift:=xlUp

'Insere template completo
template.Activate

ActiveSheet.Unprotect

If FormularioDados.Betim1310.Value = False Then
    Range("M1:Q24").Select
    Selection.Copy
Else
    Range("M25:Q61").Select
    Selection.Copy
End If

resultado.Activate
Range("A3").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False

template.Activate
ActiveSheet.Protect

'Formata tabela de tarefas
resultado.Activate
If FormularioDados.Betim1310.Value = False Then
    Range("A4:E26").Select
Else
    Range("A4:E39").Select
End If

Selection.Borders(xlDiagonalDown).LineStyle = xlNone
Selection.Borders(xlDiagonalUp).LineStyle = xlNone
With Selection.Borders(xlEdgeLeft)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlEdgeRight)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlInsideVertical)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlInsideHorizontal)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With

If FormularioDados.Betim1310.Value = False Then
    Range("A4:E26").Select
Else
    Range("A4:E39").Select
End If
With Selection
    .VerticalAlignment = xlBottom
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = False
End With

If FormularioDados.Betim1310.Value = False Then
    Range("A4:E26").Select
Else
    Range("A4:E39").Select
End If

With Selection
    .VerticalAlignment = xlBottom
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = False
End With

Range("A4:A39").Select
With Selection
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlBottom
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = False
End With
Range("C4:E39").Select
With Selection
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlBottom
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = False
End With

If FormularioDados.Betim1310.Value = False Then
    Range("A4:A26").Select
Else
    Range("A4:A39").Select
End If
Selection.NumberFormat = "0000"
If FormularioDados.Betim1310.Value = False Then
    Range("C4:C26").Select
Else
    Range("C4:C39").Select
End If

Selection.NumberFormat = "0.0"
If FormularioDados.Betim1310.Value = False Then
    Range("E4:E26").Select
Else
    Range("E4:E39").Select
End If

Selection.NumberFormat = "0.0"

If FormularioDados.chkFilho.Value = True Then
    If FormularioDados.Betim1310.Value = False Then
        Range("A24").Value = 899
        Rows("25:25").Select
        Selection.Delete Shift:=xlUp
    Else
        Range("A37").Value = 899
        Rows("38:38").Select
        Selection.Delete Shift:=xlUp
    End If
End If

HorBaciaCont = 0: HorPisoRem = 0

If FormularioDados.Betim1310.Value = False Then
    If FormularioDados.trafooleo.Value = True Then
        HorBaciaCont = 12
        HorPBSsKid = Range("E8").Value
        DurPBSskid = Range("C8").Value
        HorPBSsKid = HorPBSsKid + HorBaciaCont
        DurPBSskid = HorPBSsKid / 7.92
        Range("E8").Value = HorPBSsKid
        Range("C8").Value = Round(DurPBSskid, 1)
    End If
    
    If FormularioDados.chaparemov.Value = True Then
        HorPisoRem = 4
        HorPBSsKid = Range("E8").Value
        DurPBSskid = Range("C8").Value
        HorPBSsKid = HorPBSsKid + HorPisoRem
        DurPBSskid = HorPBSsKid / 7.92
        Range("E8").Value = HorPBSsKid
        Range("C8").Value = Round(DurPBSskid, 1)
    End If
Else
    If FormularioDados.trafooleo.Value = True Then
        HorBaciaCont = 12
        HorPBSsKid = Range("E10").Value
        DurPBSskid = Range("C10").Value
        HorPBSsKid = HorPBSsKid + HorBaciaCont
        DurPBSskid = HorPBSsKid / 7.92
        Range("E10").Value = HorPBSsKid
        Range("C10").Value = Round(DurPBSskid, 1)
    End If
    
    If FormularioDados.chaparemov.Value = True Then
        HorPisoRem = 4
        HorPBSsKid = Range("E10").Value
        DurPBSskid = Range("C10").Value
        HorPBSsKid = HorPBSsKid + HorPisoRem
        DurPBSskid = HorPBSsKid / 7.92
        Range("E10").Value = HorPBSsKid
        Range("C10").Value = Round(DurPBSskid, 1)
    End If
End If

'Elimina a tarefa de ESU quando  Betin =  false
If FormularioDados.Betim1310.Value = False And FormularioDados.tipoestrutura.Value = "Skid (mecânica)" Then
    Rows("19:19").Select
    Selection.Delete Shift:=xlUp
Else
    Rows("35:35").Select
    Selection.Delete Shift:=xlUp
End If

Range("F3").Value = "Skid (mecânica)"

'Cria tabela de totais de horas
Range("L3").Select
ActiveCell.FormulaR1C1 = "ENG"
Range("M3").Select
Selection.FormulaArray = _
    "=SUM((R4C5:R160C5)*(R4C5:R160C5<>0.1)*(R4C1:R160C1<=702)*(R4C1:R160C1<>531))"
Selection.NumberFormat = "#,##0.0"
Range("L4").Select
ActiveCell.FormulaR1C1 = "MEC"
Range("M4").Select
Selection.FormulaArray = _
        "=SUM((R4C5:R160C5)*(R4C5:R160C5<>0.1)*(R4C1:R160C1<=798)*(R4C1:R160C1>=705)*(R4C1:R160C1<>754)*(R4C1:R160C1<>755)*(R4C1:R160C1<>765)*(R4C1:R160C1<>793)*(R4C1:R160C1<>794))+SUM((R4C5:R160C5)*(R4C5:R160C5<>0.1)*(R4C1:R160C1=894))"
Selection.NumberFormat = "#,##0.0"
'Range("L5").Select
'ActiveCell.FormulaR1C1 = "ELE"
'Range("M5").Select
'Selection.FormulaArray = _
'    "=SUM((R4C5:R160C5)*(R4C5:R160C5<>0.1)*(R4C1:R160C1<=893)*(R4C1:R160C1>=799)*(R4C1:R160C1<>810)*(R4C1:R160C1<>828)*(R4C1:R160C1<>838)*(R4C1:R160C1<>858)*(R4C1:R160C1<>868))+SUM((R4C5:R160C5)*(R4C5:R160C5<>0.1)*(R4C1:R160C1=895))"
Range("L5").Select
ActiveCell.FormulaR1C1 = "TOTAL"
Range("M5").Select
Selection.FormulaArray = "=SUM(R[-3]C:R[-1]C)"
Selection.NumberFormat = "#,##0.0"
Range("L5:M5").Select
Selection.Font.Bold = True
Range("L3:M5").Select
Selection.Borders(xlDiagonalDown).LineStyle = xlNone
Selection.Borders(xlDiagonalUp).LineStyle = xlNone
With Selection.Borders(xlEdgeLeft)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlEdgeRight)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlInsideVertical)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlInsideHorizontal)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With

'Seleciona resultado
Range("A4:E4").Select
Range(Selection, Selection.End(xlDown)).Select


ActiveSheet.Protect

End Sub
```

#### `tempos_skid_eletrica` (Sub)

```vb
Sub tempos_skid_eletrica()

Application.ScreenUpdating = False

Dim HorPBAsKidElet As Double
Dim DurPBAsKidElet As Integer
Dim HorLBAsKidElet As Double
Dim DurLBAsKidElet As Integer
Dim HorDINsKidElet As Double
Dim DurDINsKidElet As Integer
Dim HorLMDsKidElet As Double
Dim DurLMDsKidElet As Integer
Dim HorPRFsKidElet As Double
Dim DurPRFsKidElet As Integer
Dim HorAFEsKidElet As Double
Dim DurAFEsKidElet As Integer
Dim PBA As Double
Dim LBA As Double
Dim DIN As Double
Dim LMD As Double
Dim PRF As Double
Dim AFE As Double

Dim HorPBSsKid As Double
Dim DurPBSskid As Integer
Dim HorBaciaCont As Integer
Dim HorPisoRem As Integer

If FormularioDados.txtPEP.Value <> "" Then
    ThisWorkbook.Application.StatusBar = "PEP " & FormularioDados.txtPEP.Value & ": Calculando tempos"
Else
    ThisWorkbook.Application.StatusBar = "Calculando tempos"
End If

ActiveSheet.Unprotect

'Limpa aba de resultados
Rows("4:245").Select
Selection.Delete Shift:=xlUp

'Insere template completo
template.Activate

If FormularioDados.Betim1310.Value = True Then

    ActiveSheet.Unprotect
    Range("AE1:AI69").Select
    Selection.Copy
    
    resultado.Activate
    Range("A3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    template.Activate
    ActiveSheet.Protect
    
    'Formata tabela de tarefas
    resultado.Activate
    Range("A4:E71").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    Range("C4:E71").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("A4:A71").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.NumberFormat = "0000"
    Range("E4:E71").Select
    Selection.NumberFormat = "0.0"
    Range("C4:C71").Select
    Selection.NumberFormat = "0.0"
    
    If FormularioDados.chkFilho.Value = True Then
        Range("A69").Value = 899
        Rows("70:70").Select
        Selection.Delete Shift:=xlUp
    End If
    
    If FormularioDados.SemEng.Value = True Then
        Range("A4:E37").Select
        Selection.Delete Shift:=xlUp
    End If

Else

    ActiveSheet.Unprotect
    Range("AK1:AO69").Select
    Selection.Copy
    
    resultado.Activate
    Range("A3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    template.Activate
    ActiveSheet.Protect
    
    'Formata tabela de tarefas
    resultado.Activate
    Range("A4:E71").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    Range("C4:E71").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("A4:A71").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.NumberFormat = "0000"
    Range("E4:E71").Select
    Selection.NumberFormat = "0.0"
    Range("C4:C71").Select
    Selection.NumberFormat = "0.0"
    
    If FormularioDados.chkFilho.Value = True Then
        Range("A69").Value = 899
        Rows("70:70").Select
        Selection.Delete Shift:=xlUp
    End If
    
    If FormularioDados.SemEng.Value = True Then
        Range("A4:E37").Select
        Selection.Delete Shift:=xlUp
    End If

End If

HorBaciaCont = 0: HorPisoRem = 0

If FormularioDados.trafooleo.Value = True Then
    HorBaciaCont = 12
    HorPBSsKid = Range("E10").Value
    DurPBSskid = Range("C10").Value
    HorPBSsKid = HorPBSsKid + HorBaciaCont
    DurPBSskid = HorPBSsKid / 7.92
    Range("E10").Value = HorPBSsKid
    Range("C10").Value = Round(DurPBSskid, 1)
End If

If FormularioDados.chaparemov.Value = True Then
    HorPisoRem = 4
    HorPBSsKid = Range("E10").Value
    DurPBSskid = Range("C10").Value
    HorPBSsKid = HorPBSsKid + HorPisoRem
    DurPBSskid = HorPBSsKid / 7.92
    Range("E10").Value = HorPBSsKid
    Range("C10").Value = Round(DurPBSskid, 1)
End If

PBA = 0: LBA = 0: DIN = 0: LMD = 0: PRF = 0: AFE = 0

If FormularioDados.nrcolunas > 0 And FormularioDados.SemEng.Value = False Then
    PBA = 8: LBA = 2.5: DIN = 8: LMD = 2.5: PRF = 6: AFE = 8
    'Horas DIN
    HorDINsKidElet = Range("E27").Value
    DurDINsKidElet = 4
    HorDINsKidElet = DIN
    Range("E27").Value = HorDINsKidElet
    Range("C27").Value = Round(DurDINsKidElet, 1)
    
    'Horas LMD
    HorLMDsKidElet = Range("E29").Value
    DurLMDsKidElet = 1
    HorLMDsKidElet = LMD
    Range("E29").Value = HorLMDsKidElet
    Range("C29").Value = Round(DurLMDsKidElet, 1)
    
    'Horas PBA
    HorPBAsKidElet = Range("E31").Value
    DurPBAsKidElet = 2
    HorPBAsKidElet = PBA
    Range("E31").Value = HorPBAsKidElet
    Range("C31").Value = Round(DurPBAsKidElet, 1)
    
    'Horas LBA
    HorLBAsKidElet = Range("E33").Value
    DurLBAsKidElet = 2
    HorLBAsKidElet = LBA
    Range("E33").Value = HorLBAsKidElet
    Range("C33").Value = Round(DurLBAsKidElet, 1)
   
    'Horas PRF
    HorPRFsKidElet = Range("E35").Value
    DurPRFsKidElet = 15
    HorPRFsKidElet = PRF
    Range("E35").Value = HorPRFsKidElet
    Range("C35").Value = Round(DurPRFsKidElet, 1)
    
    'Horas AFE
    HorAFEsKidElet = Range("E36").Value
    DurAFEsKidElet = 10
    HorAFEsKidElet = AFE
    Range("E36").Value = HorAFEsKidElet
    Range("C36").Value = Round(DurAFEsKidElet, 1)
End If

Range("F3").Value = "SKID"

'Elimina a tarefa IID quando for ITA
If FormularioDados.Betim1310.Value = False Then
    Rows("60:60").Select
    Selection.Delete Shift:=xlUp
End If

'Elimina a tarefa de ESU quando  Betin =  false
If FormularioDados.Betim1310.Value = False Then
    Rows("40:40").Select
    Selection.Delete Shift:=xlUp
Else
    Rows("42:42").Select
    Selection.Delete Shift:=xlUp
End If

'Cria tabela de totais de horas
Range("L3").Select
ActiveCell.FormulaR1C1 = "ENG"
Range("M3").Select
Selection.FormulaArray = _
    "=SUM((R4C5:R160C5)*(R4C5:R160C5<>0.1)*(R4C1:R160C1<=702)*(R4C1:R160C1<>531)*(R4C1:R160C1<>581)*(R4C1:R160C1<>585)*(R4C1:R160C1<>589))"
Selection.NumberFormat = "#,##0.0"
Range("L4").Select
ActiveCell.FormulaR1C1 = "MEC"
Range("M4").Select
Selection.FormulaArray = _
    "=SUM((R4C5:R160C5)*(R4C5:R160C5<>0.1)*(R4C1:R160C1<=798)*(R4C1:R160C1>=705)*(R4C1:R160C1<>754)*(R4C1:R160C1<>755)*(R4C1:R160C1<>765)*(R4C1:R160C1<>793)*(R4C1:R160C1<>794))+SUM((R4C5:R160C5)*(R4C5:R160C5<>0.1)*(R4C1:R160C1=894))"
Selection.NumberFormat = "#,##0.0"
Range("L5").Select
ActiveCell.FormulaR1C1 = "ELE"
Range("M5").Select
Selection.FormulaArray = _
    "=SUM((R4C5:R160C5)*(R4C5:R160C5<>0.1)*(R4C1:R160C1<=893)*(R4C1:R160C1>=799)*(R4C1:R160C1<>810)*(R4C1:R160C1<>828)*(R4C1:R160C1<>838)*(R4C1:R160C1<>858)*(R4C1:R160C1<>868))+SUM((R4C5:R160C5)*(R4C5:R160C5<>0.1)*(R4C1:R160C1=895))"
Selection.NumberFormat = "#,##0.0"
Range("L6").Select
ActiveCell.FormulaR1C1 = "TOTAL"
Range("M6").Select
Selection.FormulaArray = "=SUM(R[-3]C:R[-1]C)"
Selection.NumberFormat = "#,##0.0"
Range("L6:M6").Select
Selection.Font.Bold = True
Range("L3:M6").Select
Selection.Borders(xlDiagonalDown).LineStyle = xlNone
Selection.Borders(xlDiagonalUp).LineStyle = xlNone
With Selection.Borders(xlEdgeLeft)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlEdgeRight)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlInsideVertical)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlInsideHorizontal)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With

'Seleciona resultado
Range("A4:E4").Select
Range(Selection, Selection.End(xlDown)).Select

ActiveSheet.Protect

Application.ScreenUpdating = True

End Sub
```

#### `tempos_ESSW_eletrica` (Sub)

```vb
Sub tempos_ESSW_eletrica()

Application.ScreenUpdating = False

If FormularioDados.txtPEP.Value <> "" Then
    ThisWorkbook.Application.StatusBar = "PEP " & FormularioDados.txtPEP.Value & ": Calculando tempos"
Else
    ThisWorkbook.Application.StatusBar = "Calculando tempos"
End If

ActiveSheet.Unprotect

'Limpa aba de resultados
Rows("4:245").Select
Selection.Delete Shift:=xlUp

'Insere template completo
template.Activate

ActiveSheet.Unprotect
Range("S1:W15").Select
Selection.Copy

resultado.Activate
Range("A3").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False

template.Activate
ActiveSheet.Protect

'Formata tabela de tarefas
resultado.Activate
Range("A4:E17").Select
Selection.Borders(xlDiagonalDown).LineStyle = xlNone
Selection.Borders(xlDiagonalUp).LineStyle = xlNone
With Selection.Borders(xlEdgeLeft)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlEdgeRight)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlInsideVertical)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlInsideHorizontal)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With

Range("C4:E17").Select
With Selection
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlBottom
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = False
End With
Range("A4:A17").Select
With Selection
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlBottom
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = False
End With
Selection.NumberFormat = "0000"
Range("E4:E17").Select
Selection.NumberFormat = "0.0"
Range("C4:C17").Select
Selection.NumberFormat = "0.0"

If FormularioDados.chkFilho.Value = True Then
    Range("A15").Value = 899
    Rows("16:16").Select
    Selection.Delete Shift:=xlUp
End If

Range("F3").Value = "ESSW"

'Cria tabela de totais de horas
Range("L3").Select
ActiveCell.FormulaR1C1 = "ENG"
Range("M3").Select
Selection.FormulaArray = _
    "=SUM((R4C5:R160C5)*(R4C5:R160C5<>0.1)*(R4C1:R160C1<=702))"
Selection.NumberFormat = "#,##0.0"
'Range("L4").Select
'ActiveCell.FormulaR1C1 = "MEC"
'Range("M4").Select
'Selection.FormulaArray = _
'    "=SUM((R4C5:R160C5)*(R4C5:R160C5<>0.1)*(R4C1:R160C1<=798)*(R4C1:R160C1>=705)*(R4C1:R160C1<>754)*(R4C1:R160C1<>755)*(R4C1:R160C1<>765)*(R4C1:R160C1<>793)*(R4C1:R160C1<>794))+SUM((R4C5:R160C5)*(R4C5:R160C5<>0.1)*(R4C1:R160C1=894))"
'Range("L5").Select
'ActiveCell.FormulaR1C1 = "ELE"
'Range("M5").Select
'Selection.FormulaArray = _
'    "=SUM((R4C5:R160C5)*(R4C5:R160C5<>0.1)*(R4C1:R160C1<=893)*(R4C1:R160C1>=799)*(R4C1:R160C1<>810)*(R4C1:R160C1<>828)*(R4C1:R160C1<>838)*(R4C1:R160C1<>858)*(R4C1:R160C1<>868))+SUM((R4C5:R160C5)*(R4C5:R160C5<>0.1)*(R4C1:R160C1=895))"
Range("L4").Select
ActiveCell.FormulaR1C1 = "TOTAL"
Range("M4").Select
Selection.FormulaArray = "=SUM(R[-3]C:R[-1]C)"
Selection.NumberFormat = "#,##0.0"
Range("L4:M4").Select
Selection.Font.Bold = True
Range("L3:M4").Select
Selection.Borders(xlDiagonalDown).LineStyle = xlNone
Selection.Borders(xlDiagonalUp).LineStyle = xlNone
With Selection.Borders(xlEdgeLeft)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlEdgeRight)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlInsideVertical)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlInsideHorizontal)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With

'Seleciona resultado
Range("A4:E4").Select
Range(Selection, Selection.End(xlDown)).Select

ActiveSheet.Protect

Application.ScreenUpdating = True


End Sub
```

#### `tempos_pilotis` (Sub)

```vb
Sub tempos_pilotis()

Application.ScreenUpdating = False

If FormularioDados.txtPEP.Value <> "" Then
    ThisWorkbook.Application.StatusBar = "PEP " & FormularioDados.txtPEP.Value & ": Calculando tempos"
Else
    ThisWorkbook.Application.StatusBar = "Calculando tempos"
End If

ActiveSheet.Unprotect

'Limpa aba de resultados
Rows("4:245").Select
Selection.Delete Shift:=xlUp

'Insere template completo
template.Activate

ActiveSheet.Unprotect
Range("Y1:AC8").Select
Selection.Copy

resultado.Activate
Range("A3").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False

template.Activate
ActiveSheet.Protect

'Formata tabela de tarefas
resultado.Activate
Range("A4:E10").Select
Selection.Borders(xlDiagonalDown).LineStyle = xlNone
Selection.Borders(xlDiagonalUp).LineStyle = xlNone
With Selection.Borders(xlEdgeLeft)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlEdgeRight)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlInsideVertical)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlInsideHorizontal)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With

Range("C4:E10").Select
With Selection
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlBottom
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = False
End With
Range("A4:A10").Select
With Selection
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlBottom
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = False
End With
Selection.NumberFormat = "0000"
Range("E4:E10").Select
Selection.NumberFormat = "0.0"
Range("C4:C10").Select
Selection.NumberFormat = "0.0"

If FormularioDados.chkFilho.Value = True Then
    Range("A8").Value = 899
    Rows("9:9").Select
    Selection.Delete Shift:=xlUp
End If

Range("F3").Value = "PILOTIS"

'Cria tabela de totais de horas
Range("L3").Select
ActiveCell.FormulaR1C1 = "ENG"
Range("M3").Select
Selection.FormulaArray = _
    "=SUM((R4C5:R160C5)*(R4C5:R160C5<>0.1)*(R4C1:R160C1<=702))"
Selection.NumberFormat = "#,##0.0"
''Range("L4").Select
''ActiveCell.FormulaR1C1 = "MEC"
''Range("M4").Select
''Selection.FormulaArray = _
''    "=SUM((R4C5:R160C5)*(R4C5:R160C5<>0.1)*(R4C1:R160C1<=798)*(R4C1:R160C1>=705)*(R4C1:R160C1<>754)*(R4C1:R160C1<>755)*(R4C1:R160C1<>765)*(R4C1:R160C1<>793)*(R4C1:R160C1<>794))+SUM((R4C5:R160C5)*(R4C5:R160C5<>0.1)*(R4C1:R160C1=894))"
''Range("L5").Select
''ActiveCell.FormulaR1C1 = "ELE"
''Range("M5").Select
''Selection.FormulaArray = _
''    "=SUM((R4C5:R160C5)*(R4C5:R160C5<>0.1)*(R4C1:R160C1<=893)*(R4C1:R160C1>=799)*(R4C1:R160C1<>810)*(R4C1:R160C1<>828)*(R4C1:R160C1<>838)*(R4C1:R160C1<>858)*(R4C1:R160C1<>868))+SUM((R4C5:R160C5)*(R4C5:R160C5<>0.1)*(R4C1:R160C1=895))"
Range("L4").Select
ActiveCell.FormulaR1C1 = "TOTAL"
Range("M4").Select
Selection.FormulaArray = "=SUM(R[-3]C:R[-1]C)"
Selection.NumberFormat = "#,##0.0"
Range("L4:M4").Select
Selection.Font.Bold = True
Range("L3:M4").Select
Selection.Borders(xlDiagonalDown).LineStyle = xlNone
Selection.Borders(xlDiagonalUp).LineStyle = xlNone
With Selection.Borders(xlEdgeLeft)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlEdgeRight)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlInsideVertical)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlInsideHorizontal)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With

'Seleciona resultado
Range("A4:E4").Select
Range(Selection, Selection.End(xlDown)).Select

ActiveSheet.Protect

Application.ScreenUpdating = True

End Sub
```

#### `definir_template` (Sub)

```vb
Sub definir_template()

Dim template_eng_mec As String
Dim altern_eng_mec As String
Dim template_eng_ele As String
Dim altern_eng_ele As String

Dim template_mec1 As String
Dim altern_mec1 As String
Dim template_mec2 As String
Dim altern_mec2 As String
Dim template_mec3 As String
Dim altern_mec3 As String
Dim template_mec4 As String
Dim altern_mec4 As String

Dim template_mec5 As String
Dim altern_mec5 As String
Dim template_mec6 As String
Dim altern_mec6 As String
Dim template_mec7 As String
Dim altern_mec7 As String
Dim template_mec8 As String
Dim altern_mec8 As String

Dim template_acess As String
Dim altern_acess As String
Dim template_eletr As String
Dim altern_eletr As String

If FormularioDados.txtPEP.Value <> "" Then
    ThisWorkbook.Application.StatusBar = "PEP " & FormularioDados.txtPEP.Value & ": Definindo template"
Else
    ThisWorkbook.Application.StatusBar = "Definindo template"
End If

'Selecionar template
seletor_template.Activate

If FormularioDados.Betim1310.Value = False Then
    If FormularioDados.tipoestrutura.Value = "Container Solar" Or FormularioDados.tipoestrutura.Value = "Skid (mecânica)" Or FormularioDados.tipoestrutura.Value = "ESSW (mecânica)" Or FormularioDados.tipoestrutura.Value = "ESSW (elétrica)" Or FormularioDados.tipoestrutura.Value = "Serviço Engenharia" Then
        ActiveSheet.Range("$A$1:$V$108").AutoFilter Field:=1, Criteria1:=FormularioDados.tipoestrutura.Value
    Else
        ActiveSheet.Range("$A$1:$V$108").AutoFilter Field:=1, Criteria1:="Eletrocentro"
    End If
End If

If FormularioDados.tipoestrutura.Value = "Skid (com elétrica)" And FormularioDados.Betim1310.Value = True Then
    ActiveSheet.Range("$A$1:$V$108").AutoFilter Field:=1, Criteria1:="Skid (com elétrica)B"
End If

If FormularioDados.tipoestrutura.Value = "Pilotis" Then
    If FormularioDados.Betim1310.Value = True Then
        ActiveSheet.Range("$A$1:$V$108").AutoFilter Field:=1, Criteria1:="PilotisB"
    Else
        ActiveSheet.Range("$A$1:$V$108").AutoFilter Field:=1, Criteria1:="Pilotis"
    End If
End If

If FormularioDados.tipoestrutura.Value = "Container Solar" And FormularioDados.Betim1310.Value = True Then

    If FormularioDados.SemEng.Value = False Then
        ActiveSheet.Range("$A$1:$V$108").AutoFilter Field:=1, Criteria1:="Container SolarBI"
    Else
        ActiveSheet.Range("$A$1:$V$108").AutoFilter Field:=1, Criteria1:="Container SolarBII"
    End If
    
End If

If FormularioDados.tipoestrutura.Value <> "Container Solar" And FormularioDados.tipoestrutura.Value <> "Skid (mecânica)" And FormularioDados.tipoestrutura.Value <> "Pilotis" And FormularioDados.tipoestrutura.Value <> "ESSW (mecânica)" And FormularioDados.tipoestrutura.Value <> "ESSW (elétrica)" And FormularioDados.Betim1310.Value = True Then
    ActiveSheet.Range("$A$1:$V$108").AutoFilter Field:=1, Criteria1:="EletrocentroB"
End If

'##TESTE
If FormularioDados.nrmodulos.Value = "1 Módulo" Or FormularioDados.tipoestrutura.Value = "Serviço Engenharia" Then
    ActiveSheet.Range("$A$1:$V$108").AutoFilter Field:=2, Criteria1:=1
    If FormularioDados.tipoestrutura.Value = "Serviço Engenharia" Then GoTo Pular
End If
If FormularioDados.nrmodulos.Value = "2 Módulos" Then ActiveSheet.Range("$A$1:$V$108").AutoFilter Field:=2, Criteria1:=2
If FormularioDados.nrmodulos.Value = "3 Módulos" Then ActiveSheet.Range("$A$1:$V$108").AutoFilter Field:=2, Criteria1:=3
If FormularioDados.nrmodulos.Value = "4 Módulos" Then ActiveSheet.Range("$A$1:$V$108").AutoFilter Field:=2, Criteria1:=4
If FormularioDados.nrmodulos.Value = "5 Módulos" Then ActiveSheet.Range("$A$1:$V$108").AutoFilter Field:=2, Criteria1:=5
If FormularioDados.nrmodulos.Value = "6 Módulos" Then ActiveSheet.Range("$A$1:$V$108").AutoFilter Field:=2, Criteria1:=6
If FormularioDados.nrmodulos.Value = "7 Módulos" Then ActiveSheet.Range("$A$1:$V$108").AutoFilter Field:=2, Criteria1:=7
If FormularioDados.nrmodulos.Value = "8 Módulos" Then ActiveSheet.Range("$A$1:$V$108").AutoFilter Field:=2, Criteria1:=8

If FormularioDados.tipomaq.Value = "Roof Top" Then
    ActiveSheet.Range("$A$1:$V$108").AutoFilter Field:=3, Criteria1:="Sim"
Else
    ActiveSheet.Range("$A$1:$V$108").AutoFilter Field:=3, Criteria1:="Não"
End If

If FormularioDados.Betim1310.Value = True Then
    ActiveSheet.Range("$A$1:$V$108").AutoFilter Field:=4, Criteria1:="-"
ElseIf FormularioDados.seguranca.Value <> "Não possui" And FormularioDados.seguranca.Value <> "Não aplicável" Then
    ActiveSheet.Range("$A$1:$V$108").AutoFilter Field:=4, Criteria1:="Sim"
Else
    ActiveSheet.Range("$A$1:$V$108").AutoFilter Field:=4, Criteria1:="Não"
End If

If FormularioDados.testesw.Value = True Then
    ActiveSheet.Range("$A$1:$V$108").AutoFilter Field:=5, Criteria1:="Sim"
Else
    ActiveSheet.Range("$A$1:$V$108").AutoFilter Field:=5, Criteria1:="Não"
End If

If FormularioDados.tipoestrutura.Value = "Skid (com elétrica)" And FormularioDados.Betim1310.Value = True Then
    ActiveSheet.Range("$A$1:$V$108").AutoFilter Field:=1, Criteria1:="Skid (com elétrica)B"
End If

If FormularioDados.tipoestrutura.Value = "Skid (mecânica)" And FormularioDados.Betim1310.Value = True Then
    ActiveSheet.Range("$A$1:$V$108").AutoFilter Field:=1, Criteria1:="Skid (mecânica)B"
End If

Pular:
'Capturar template selecionado
ActiveCell.SpecialCells(xlLastCell).Select
If Range("A1").Value = 1 Then
    template_eng_mec = Cells(ActiveCell.Row, 7).Value
    altern_eng_mec = Cells(ActiveCell.Row, 8).Value
    
    template_eng_ele = Cells(ActiveCell.Row, 9).Value
    altern_eng_ele = Cells(ActiveCell.Row, 10).Value
    
    template_mec1 = Cells(ActiveCell.Row, 11).Value
    altern_mec1 = Cells(ActiveCell.Row, 12).Value
    
    template_mec2 = Cells(ActiveCell.Row, 13).Value
    altern_mec2 = Cells(ActiveCell.Row, 14).Value
    
    template_mec3 = Cells(ActiveCell.Row, 15).Value
    altern_mec3 = Cells(ActiveCell.Row, 16).Value
    
    template_mec4 = Cells(ActiveCell.Row, 17).Value
    altern_mec4 = Cells(ActiveCell.Row, 18).Value
    
    template_mec5 = Cells(ActiveCell.Row, 19).Value
    altern_mec5 = Cells(ActiveCell.Row, 20).Value
    
    template_mec6 = Cells(ActiveCell.Row, 21).Value
    altern_mec6 = Cells(ActiveCell.Row, 22).Value
    
    template_mec7 = Cells(ActiveCell.Row, 23).Value
    altern_mec7 = Cells(ActiveCell.Row, 24).Value
    
    template_mec8 = Cells(ActiveCell.Row, 25).Value
    altern_mec8 = Cells(ActiveCell.Row, 26).Value
    
    template_acess = Cells(ActiveCell.Row, 27).Value
    altern_acess = Cells(ActiveCell.Row, 28).Value
    
    template_eletr = Cells(ActiveCell.Row, 29).Value
    altern_eletr = Cells(ActiveCell.Row, 30).Value
        
Else
    template_eng_mec = "???"
    altern_eng_mec = "?"
    template_eng_ele = "???"
    altern_eng_ele = "?"
    template_mec1 = "???"
    altern_mec1 = "?"
    template_mec2 = "???"
    altern_mec2 = "?"
    template_mec3 = "???"
    altern_mec3 = "?"
    template_mec4 = "???"
    altern_mec4 = "?"
    template_mec5 = "???"
    altern_mec5 = "?"
    template_mec6 = "???"
    altern_mec6 = "?"
    template_mec7 = "???"
    altern_mec7 = "?"
    template_mec8 = "???"
    altern_mec8 = "?"
    template_acess = "???"
    altern_acess = "?"
    template_eletr = "???"
    altern_eletr = "?"
End If

Rows("1:1").Select
Selection.AutoFilter
Selection.AutoFilter

resultado.Activate
ActiveSheet.Unprotect

Range("H3:J14").Select
Selection.Borders(xlDiagonalDown).LineStyle = xlNone
Selection.Borders(xlDiagonalUp).LineStyle = xlNone
With Selection.Borders(xlEdgeLeft)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlEdgeRight)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlInsideVertical)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlInsideHorizontal)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With

Range("H4").Select
ActiveCell.FormulaR1C1 = "ENG ELE"
Range("H5").Select
ActiveCell.FormulaR1C1 = "MEC 1"
Range("H6").Select
ActiveCell.FormulaR1C1 = "MEC 2"
Range("H7").Select
ActiveCell.FormulaR1C1 = "MEC 3"
Range("H8").Select
ActiveCell.FormulaR1C1 = "MEC 4"
Range("H9").Select
ActiveCell.FormulaR1C1 = "MEC 5"
Range("H10").Select
ActiveCell.FormulaR1C1 = "MEC 6"
Range("H11").Select
ActiveCell.FormulaR1C1 = "MEC 7"
Range("H12").Select
ActiveCell.FormulaR1C1 = "MEC 8"
Range("H13").Select
ActiveCell.FormulaR1C1 = "ACESS"
Range("H14").Select
ActiveCell.FormulaR1C1 = "ELETR"

Range("I3").Value = template_eng_mec
Range("J3").Value = altern_eng_mec
Range("I4").Value = template_eng_ele
Range("J4").Value = altern_eng_ele
Range("I5").Value = template_mec1
Range("J5").Value = altern_mec1
Range("I6").Value = template_mec2
Range("J6").Value = altern_mec2
Range("I7").Value = template_mec3
Range("J7").Value = altern_mec3
Range("I8").Value = template_mec4
Range("J8").Value = altern_mec4
Range("I9").Value = template_mec5
Range("J9").Value = altern_mec5
Range("I10").Value = template_mec6
Range("J10").Value = altern_mec6
Range("I11").Value = template_mec7
Range("J11").Value = altern_mec7
Range("I12").Value = template_mec8
Range("J12").Value = altern_mec8
Range("I13").Value = template_acess
Range("J13").Value = altern_acess
Range("I14").Value = template_eletr
Range("J14").Value = altern_eletr

If FormularioDados.SemEng.Value = True And FormularioDados.tipoestrutura.Value = "Container Solar" Then
   
   Range("I3").Value = "-"
   Range("J3").Value = "-"
   Range("I4").Value = "-"
   Range("J4").Value = "-"
   
   End If


ActiveSheet.Protect

End Sub
```

#### `criar_DRs` (Sub)

```vb
Sub criar_DRs()

Dim SapGuiAuto As Object
Dim Application As Object
Dim Connection As Object
Dim Session As Object
Dim WScript As Object
Dim j As Long
Dim sessao_ok As Boolean

'Conexão com o Objeto SAP
Set SapGuiAuto = GetObject("SAPGUI")
Set Application = SapGuiAuto.GetScriptingEngine
Set Connection = Application.Children(0)


sessao_ok = False
For j = 0 To Application.Children(0).Sessions.Count() - 1
    Set Session = Connection.Children(CLng(j))
    If Session.ActiveWindow.Text = "SAP Easy Access" Then
        sessao_ok = True
        Exit For
    End If
Next

If sessao_ok = False Then
    MsgBox "Nenhuma janela do SAP na tela inicial foi encontrada. Programa interrompido.", vbOKOnly
    Exit Sub
End If

ThisWorkbook.Application.StatusBar = "PEP " & FormularioDados.txtPEP.Value & ": Criando diagramas de rede"

If Range("I3").Value = "???" Or Range("I4").Value = "???" Or Range("I5").Value = "???" Or Range("I6").Value = "???" Or Range("I7").Value = "???" Or Range("I8").Value = "???" Or Range("I9").Value = "???" Or Range("I10").Value = "???" Then
    MsgBox "Template não encontrado. Impossível criar Diagramas de Rede.", vbOKOnly
    Exit Sub
End If

'Abrir CJ20N e inserir PEP
Session.findById("wnd[0]").maximize
Session.findById("wnd[0]/tbar[0]/okcd").Text = "/NCJ20N"
Session.findById("wnd[0]").sendVKey 0
Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[0]/shell").pressButton "OPEN"
Session.findById("wnd[1]/usr/ctxtCNPB_W_ADD_OBJ_DYN-PROJ_EXT").Text = ""
Session.findById("wnd[1]/usr/ctxtCNPB_W_ADD_OBJ_DYN-PRPS_EXT").Text = FormularioDados.txtPEP.Value
Session.findById("wnd[1]").sendVKey 0


'Encerrar programa em caso de mensagem de erro (PEP inexistente)
If Session.ActiveWindow.Text = "Erro" Then
    Session.findById("wnd[2]").sendVKey 0
    Session.findById("wnd[1]").Close
    Session.findById("wnd[0]/tbar[0]/btn[15]").press
    MsgBox "PEP informado não existe. Impossível criar Diagramas de Rede.", vbOKOnly
    Exit Sub
End If

'Por em modo de edição
If Session.ActiveWindow.Text = "Project Builder: exibir subprojeto " & FormularioDados.txtPEP.Value Then
    Session.findById("wnd[0]/tbar[1]/btn[13]").press
    If Session.findById("wnd[0]/sbar").Text = "Não foram bloqueados todos os objetos (ver protocolo de bloqueio)" Then
        Session.findById("wnd[0]/tbar[0]/btn[15]").press
        MsgBox "PEP está aberto por outro usuário. Verificar e tentar novamente mais tarde.", vbOKOnly
        Exit Sub
    End If
Else
    If Session.findById("wnd[0]/sbar").Text = "Não foram bloqueados todos os objetos (ver protocolo de bloqueio)" Then
        Session.findById("wnd[0]/tbar[0]/btn[15]").press
        MsgBox "PEP está aberto por outro usuário. Verificar e tentar novamente mais tarde.", vbOKOnly
        Exit Sub
    End If
End If

'Verificar se já há diagramas criados
Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "NETW_OVW"
If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtNETW_OVW-AUFNR[0,0]").Text <> "" Then
    Session.findById("wnd[0]/tbar[0]/btn[15]").press
    MsgBox "Já existem diagramas de rede criados neste PEP. Criação de novos DRs cancelada.", vbOKOnly
    Exit Sub
End If

'Verificar parâmetros da definição do projeto
Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000001"
Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCJWB:3998/tabsPTABSCR/tabpPCNT").Select
If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCJWB:3998/tabsPTABSCR/tabpPCNT/ssubSUBSCR2:SAPLCJWB:1404/cmbPROJ-SCPRF").Key <> "Z00000000002" Then
    Session.findById("wnd[0]/tbar[0]/btn[15]").press
    MsgBox "Perfil de programação da Definição do projeto incorreto. Favor verificar.", vbOKOnly
    Exit Sub
End If
If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCJWB:3998/tabsPTABSCR/tabpPCNT/ssubSUBSCR2:SAPLCJWB:1404/cmbPROJ-SCHTYP").Key <> "" Then
    Session.findById("wnd[0]/tbar[0]/btn[15]").press
    MsgBox "Cenário de programação de prazo da Definição do projeto incorreto. Favor verificar.", vbOKOnly
    Exit Sub
End If
If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCJWB:3998/tabsPTABSCR/tabpPCNT/ssubSUBSCR2:SAPLCJWB:1404/cmbPROJ-VGPLF").Key <> "3" Then
    Session.findById("wnd[0]/tbar[0]/btn[15]").press
    MsgBox "Método de planejamento de datas base da Definição do projeto incorreto. Favor verificar.", vbOKOnly
    Exit Sub
End If
If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCJWB:3998/tabsPTABSCR/tabpPCNT/ssubSUBSCR2:SAPLCJWB:1404/cmbPROJ-EWPLF").Key <> "3" Then
    Session.findById("wnd[0]/tbar[0]/btn[15]").press
    MsgBox "Método de planejamento de datas previstas da Definição do projeto incorreto. Favor verificar.", vbOKOnly
    Exit Sub
End If

'Quando for solar troca de centro de custo responsável: de 30111174 para 31011174
Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode "000002"
Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_DETAIL/shellcont/shell").pressButton "WBSE_DET"
If FormularioDados.chkSolar.Value = True Then
   Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCJWB:3999/tabsTABCJWB/tabpGRND/ssubSUBSCR1:SAPLCJWB:1210/ctxtPRPS-FKSTL").Text = "31011174"
   Session.findById("wnd[0]").sendVKey 0
End If
'Verificar se o centro do PEP está correto
Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCJWB:3999/tabsTABCJWB/tabpORGA").Select
If FormularioDados.tipoestrutura.Value <> "ESSW (elétrica)" And Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCJWB:3999/tabsTABCJWB/tabpORGA/ssubSUBSCR1:SAPLCJWB:1410/ctxtPRPS-WERKS").Text <> "1306" And FormularioDados.Betim1310.Value = False Then
    Session.findById("wnd[0]/tbar[0]/btn[15]").press
    MsgBox "Centro do PEP no SAP incorreto. Favor verificar. Criação de novos DRs cancelada.", vbOKOnly
    Exit Sub
End If
If FormularioDados.tipoestrutura.Value = "ESSW (elétrica)" And Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCJWB:3999/tabsTABCJWB/tabpORGA/ssubSUBSCR1:SAPLCJWB:1410/ctxtPRPS-WERKS").Text <> "1320" Then
    Session.findById("wnd[0]/tbar[0]/btn[15]").press
    MsgBox "Centro do PEP no SAP incorreto. Favor verificar. Criação de novos DRs cancelada.", vbOKOnly
    Exit Sub
End If
If FormularioDados.Betim1310.Value = False And Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCJWB:3999/tabsTABCJWB/tabpORGA/ssubSUBSCR1:SAPLCJWB:1410/ctxtPRPS-WERKS").Text <> "1306" Then
    Session.findById("wnd[0]/tbar[0]/btn[15]").press
    MsgBox "Centro do PEP no SAP incorreto. Favor verificar. Criação de novos DRs cancelada.", vbOKOnly
    Exit Sub
End If
If FormularioDados.Betim1310.Value = True And Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCJWB:3999/tabsTABCJWB/tabpORGA/ssubSUBSCR1:SAPLCJWB:1410/ctxtPRPS-WERKS").Text <> "1313" Then
    Session.findById("wnd[0]/tbar[0]/btn[15]").press
    MsgBox "Centro do PEP no SAP incorreto. Favor verificar. Criação de novos DRs cancelada.", vbOKOnly
    Exit Sub
End If

'Inserir diagrama engenharia mecânica
If FormularioDados.tipoestrutura.Value <> "ESSW (elétrica)" And FormularioDados.SemEng.Value = False Then
    Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode "000002"
    Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").nodeContextMenu "000002"
    Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectContextMenuItem "COPY_NET"
    Session.findById("wnd[1]/usr/tabsTABSTR_2160/tabpTAB_TEMPLATE/ssubSUBSCREEN_2160:SAPLCOKO:2161/ctxtCAUFVD-STDNR").Text = Range("I3").Value
    Session.findById("wnd[1]/usr/tabsTABSTR_2160/tabpTAB_TEMPLATE/ssubSUBSCREEN_2160:SAPLCOKO:2161/txtCAUFVD-PLNAL").Text = Range("J3").Value
    Session.findById("wnd[1]").sendVKey 0
    If Session.findById("wnd[0]/sbar").Text = "Indicar um tipo de diagrama rede" Then
        Session.findById("wnd[1]/usr/tabsTABSTR_2160/tabpTAB_PARAMETER/ssubSUBSCREEN_2160:SAPLCOKO:2163/ctxtAUFPAR-PS_AUFART").Text = "ZPS1"
        Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").nodeContextMenu "000002"
        Session.findById("wnd[1]").sendVKey 0
    End If
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subIDENTIFICATION:SAPLCOKO:2816/txtCAUFVD-KTEXT").Text = "Engenharia Mecânica"
    Session.findById("wnd[0]").sendVKey 0
    If Session.ActiveWindow.Text = "Project Builder: subprojeto " & FormularioDados.txtPEP.Value Then Session.findById("wnd[0]").sendVKey 0
    Session.findById("wnd[0]/tbar[0]/btn[3]").press
    If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
    Session.findById("wnd[0]/tbar[0]/btn[3]").press
    If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTRMN").Select
    If Session.ActiveWindow.Text = "Project Builder: subprojeto " & FormularioDados.txtPEP.Value Then
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTRMN/ssubSUBSCR_2100:SAPLCOKO:2110/ctxtCAUFVD-GSTRP").Text = ""
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTRMN/ssubSUBSCR_2100:SAPLCOKO:2110/ctxtCAUFVD-GLTRP").Text = ""
        Session.findById("wnd[0]").sendVKey 0
        If Session.findById("wnd[0]/sbar").Text = "Programação regressiva de prazos (entrar data de conclusão)" Then Session.findById("wnd[0]").sendVKey 12
        If Session.ActiveWindow.Text = "Incluir diagrama de rede padrão" Then Session.findById("wnd[1]/usr/btnSPOP-VAROPTION2").press
        Do While Session.ActiveWindow.Text = "Aceitar RD"
        Session.findById("wnd[1]/usr/btnSPOP-OPTION2").press
        Loop
        Session.findById("wnd[0]/tbar[0]/btn[3]").press
        If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
        Session.findById("wnd[0]/tbar[0]/btn[3]").press
        If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
      
    End If
    If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
    Session.findById("wnd[0]").sendVKey 0
    
    If Session.ActiveWindow.Text = "Incluir diagrama de rede padrão" Then Session.findById("wnd[1]/usr/btnSPOP-VAROPTION2").press
        Do While Session.ActiveWindow.Text = "Aceitar RD"
        Session.findById("wnd[1]/usr/btnSPOP-OPTION2").press
        Loop
    Session.findById("wnd[0]/tbar[0]/btn[3]").press
    If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
    Session.findById("wnd[0]/tbar[0]/btn[3]").press
    If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
    

End If

'Inserir diagrama engenharia elétrica

If FormularioDados.tipoestrutura.Value <> "ESSW (elétrica)" And FormularioDados.tipoestrutura.Value <> "Skid (mecânica)" And FormularioDados.tipoestrutura.Value <> "Pilotis" And FormularioDados.SemEng.Value = False Then
    Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
    Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").nodeContextMenu "000002"
    Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectContextMenuItem "COPY_NET"
    Session.findById("wnd[1]/usr/tabsTABSTR_2160/tabpTAB_TEMPLATE/ssubSUBSCREEN_2160:SAPLCOKO:2161/ctxtCAUFVD-STDNR").Text = Range("I4").Value
    Session.findById("wnd[1]/usr/tabsTABSTR_2160/tabpTAB_TEMPLATE/ssubSUBSCREEN_2160:SAPLCOKO:2161/txtCAUFVD-PLNAL").Text = Range("J4").Value
    Session.findById("wnd[1]").sendVKey 0
    If Session.findById("wnd[0]/sbar").Text = "Indicar um tipo de diagrama rede" Then
        Session.findById("wnd[1]/usr/tabsTABSTR_2160/tabpTAB_PARAMETER/ssubSUBSCREEN_2160:SAPLCOKO:2163/ctxtAUFPAR-PS_AUFART").Text = "ZPS1"
        Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").nodeContextMenu "000002"
        Session.findById("wnd[1]").sendVKey 0
    End If
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subIDENTIFICATION:SAPLCOKO:2816/txtCAUFVD-KTEXT").Text = "Engenharia Elétrica"
    Session.findById("wnd[0]").sendVKey 0
    If Session.ActiveWindow.Text = "Project Builder: subprojeto " & FormularioDados.txtPEP.Value Then Session.findById("wnd[0]").sendVKey 0
    Session.findById("wnd[0]/tbar[0]/btn[3]").press
    If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
    Session.findById("wnd[0]/tbar[0]/btn[3]").press
    If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTRMN").Select
    If Session.ActiveWindow.Text = "Project Builder: subprojeto " & FormularioDados.txtPEP.Value Then
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTRMN/ssubSUBSCR_2100:SAPLCOKO:2110/ctxtCAUFVD-GSTRP").Text = ""
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTRMN/ssubSUBSCR_2100:SAPLCOKO:2110/ctxtCAUFVD-GLTRP").Text = ""
        Session.findById("wnd[0]").sendVKey 0
        If Session.findById("wnd[0]/sbar").Text = "Programação regressiva de prazos (entrar data de conclusão)" Then Session.findById("wnd[0]").sendVKey 12
        Do While Session.ActiveWindow.Text = "Aceitar RD"
        Session.findById("wnd[1]/usr/btnSPOP-OPTION2").press
        Loop
        Session.findById("wnd[0]/tbar[0]/btn[3]").press
        If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
        Session.findById("wnd[0]/tbar[0]/btn[3]").press
        If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
    End If
    If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
    
    If Session.ActiveWindow.Text = "Incluir diagrama de rede padrão" Then Session.findById("wnd[1]/usr/btnSPOP-VAROPTION2").press

        Do While Session.ActiveWindow.Text = "Aceitar RD"
        Session.findById("wnd[1]/usr/btnSPOP-OPTION2").press
        Loop
    Session.findById("wnd[0]/tbar[0]/btn[3]").press
    If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
    Session.findById("wnd[0]/tbar[0]/btn[3]").press
    If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
End If

'Inserir diagrama módulo 1
If FormularioDados.tipoestrutura.Value <> "ESSW (elétrica)" And FormularioDados.tipoestrutura.Value <> "Pilotis" Then
    Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
    Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").nodeContextMenu "000002"
    Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectContextMenuItem "COPY_NET"
    Session.findById("wnd[1]/usr/tabsTABSTR_2160/tabpTAB_TEMPLATE/ssubSUBSCREEN_2160:SAPLCOKO:2161/ctxtCAUFVD-STDNR").Text = Range("I5").Value
    Session.findById("wnd[1]/usr/tabsTABSTR_2160/tabpTAB_TEMPLATE/ssubSUBSCREEN_2160:SAPLCOKO:2161/txtCAUFVD-PLNAL").Text = Range("J5").Value
    Session.findById("wnd[1]").sendVKey 0
    If Session.findById("wnd[0]/sbar").Text = "Indicar um tipo de diagrama rede" Then
        Session.findById("wnd[1]/usr/tabsTABSTR_2160/tabpTAB_PARAMETER/ssubSUBSCREEN_2160:SAPLCOKO:2163/ctxtAUFPAR-PS_AUFART").Text = "ZPS1"
        Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").nodeContextMenu "000002"
        Session.findById("wnd[1]").sendVKey 0
    End If
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subIDENTIFICATION:SAPLCOKO:2816/txtCAUFVD-KTEXT").Text = "Mecânica 1"
    Session.findById("wnd[0]").sendVKey 0
    Session.findById("wnd[0]").sendVKey 0
    If Session.ActiveWindow.Text = "Incluir diagrama de rede padrão" Then Session.findById("wnd[1]/usr/btnSPOP-VAROPTION2").press
    If Session.ActiveWindow.Text = "Incluir diagrama de rede padrão" Then Session.findById("wnd[1]/usr/btnSPOP-VAROPTION2").press
    
    Session.findById("wnd[0]/tbar[0]/btn[3]").press
    If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
    Session.findById("wnd[0]/tbar[0]/btn[3]").press
    If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
    
    If Session.ActiveWindow.Text = "Project Builder: subprojeto " & FormularioDados.txtPEP.Value Then Session.findById("wnd[0]").sendVKey 0
    Session.findById("wnd[0]/tbar[0]/btn[3]").press
    If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
    Session.findById("wnd[0]/tbar[0]/btn[3]").press
    If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTRMN").Select
    If Session.ActiveWindow.Text = "Project Builder: subprojeto " & FormularioDados.txtPEP.Value Then
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTRMN/ssubSUBSCR_2100:SAPLCOKO:2110/ctxtCAUFVD-GSTRP").Text = ""
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTRMN/ssubSUBSCR_2100:SAPLCOKO:2110/ctxtCAUFVD-GLTRP").Text = ""
        Session.findById("wnd[0]").sendVKey 0
        If Session.findById("wnd[0]/sbar").Text = "Programação regressiva de prazos (entrar data de conclusão)" Then Session.findById("wnd[0]").sendVKey 12
        If Session.ActiveWindow.Text = "Incluir diagrama de rede padrão" Then Session.findById("wnd[1]/usr/btnSPOP-VAROPTION2").press
        Do While Session.ActiveWindow.Text = "Aceitar RD"
        Session.findById("wnd[1]/usr/btnSPOP-OPTION2").press
        Loop
        Session.findById("wnd[0]/tbar[0]/btn[3]").press
        If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
        Session.findById("wnd[0]/tbar[0]/btn[3]").press
        If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
    End If
    If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
    
    If Session.ActiveWindow.Text = "Incluir diagrama de rede padrão" Then Session.findById("wnd[1]/usr/btnSPOP-VAROPTION2").press
    
        Do While Session.ActiveWindow.Text = "Aceitar RD"
        Session.findById("wnd[1]/usr/btnSPOP-OPTION2").press
        Loop
        Session.findById("wnd[0]/tbar[0]/btn[3]").press
        If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
        Session.findById("wnd[0]/tbar[0]/btn[3]").press
        If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
    End If

'Inserir diagrama módulo 2
If FormularioDados.nrmodulos.Value <> "1 Módulo" Then
    Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
    Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").nodeContextMenu "000002"
    Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectContextMenuItem "COPY_NET"
    Session.findById("wnd[1]/usr/tabsTABSTR_2160/tabpTAB_TEMPLATE/ssubSUBSCREEN_2160:SAPLCOKO:2161/ctxtCAUFVD-STDNR").Text = Range("I6").Value
    Session.findById("wnd[1]/usr/tabsTABSTR_2160/tabpTAB_TEMPLATE/ssubSUBSCREEN_2160:SAPLCOKO:2161/txtCAUFVD-PLNAL").Text = Range("J6").Value
    Session.findById("wnd[1]").sendVKey 0
    If Session.findById("wnd[0]/sbar").Text = "Indicar um tipo de diagrama rede" Then
        Session.findById("wnd[1]/usr/tabsTABSTR_2160/tabpTAB_PARAMETER/ssubSUBSCREEN_2160:SAPLCOKO:2163/ctxtAUFPAR-PS_AUFART").Text = "ZPS1"
        Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").nodeContextMenu "000002"
        Session.findById("wnd[1]").sendVKey 0
    End If
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subIDENTIFICATION:SAPLCOKO:2816/txtCAUFVD-KTEXT").Text = "Mecânica 2"
    Session.findById("wnd[0]").sendVKey 0
    If Session.ActiveWindow.Text = "Project Builder: subprojeto " & FormularioDados.txtPEP.Value Then Session.findById("wnd[0]").sendVKey 0
    Session.findById("wnd[0]/tbar[0]/btn[3]").press
    If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
    Session.findById("wnd[0]/tbar[0]/btn[3]").press
    If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTRMN").Select
    If Session.ActiveWindow.Text = "Project Builder: subprojeto " & FormularioDados.txtPEP.Value Then
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTRMN/ssubSUBSCR_2100:SAPLCOKO:2110/ctxtCAUFVD-GSTRP").Text = ""
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTRMN/ssubSUBSCR_2100:SAPLCOKO:2110/ctxtCAUFVD-GLTRP").Text = ""
        Session.findById("wnd[0]").sendVKey 0
        If Session.findById("wnd[0]/sbar").Text = "Programação regressiva de prazos (entrar data de conclusão)" Then Session.findById("wnd[0]").sendVKey 12
        Do While Session.ActiveWindow.Text = "Aceitar RD"
        Session.findById("wnd[1]/usr/btnSPOP-OPTION2").press
        Loop
        Session.findById("wnd[0]/tbar[0]/btn[3]").press
        If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
        Session.findById("wnd[0]/tbar[0]/btn[3]").press
        If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
    End If
    If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
    
    If Session.ActiveWindow.Text = "Incluir diagrama de rede padrão" Then Session.findById("wnd[1]/usr/btnSPOP-VAROPTION2").press
    

        Do While Session.ActiveWindow.Text = "Aceitar RD"
        Session.findById("wnd[1]/usr/btnSPOP-OPTION2").press
        Loop
    Session.findById("wnd[0]/tbar[0]/btn[3]").press
    If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
    Session.findById("wnd[0]/tbar[0]/btn[3]").press
    If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
End If

'Inserir diagrama módulo 3
If FormularioDados.nrmodulos.Value <> "1 Módulo" And FormularioDados.nrmodulos.Value <> "2 Módulos" Then
    Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
    Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").nodeContextMenu "000002"
    Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectContextMenuItem "COPY_NET"
    Session.findById("wnd[1]/usr/tabsTABSTR_2160/tabpTAB_TEMPLATE/ssubSUBSCREEN_2160:SAPLCOKO:2161/ctxtCAUFVD-STDNR").Text = Range("I7").Value
    Session.findById("wnd[1]/usr/tabsTABSTR_2160/tabpTAB_TEMPLATE/ssubSUBSCREEN_2160:SAPLCOKO:2161/txtCAUFVD-PLNAL").Text = Range("J7").Value
    Session.findById("wnd[1]").sendVKey 0
    If Session.findById("wnd[0]/sbar").Text = "Indicar um tipo de diagrama rede" Then
        Session.findById("wnd[1]/usr/tabsTABSTR_2160/tabpTAB_PARAMETER/ssubSUBSCREEN_2160:SAPLCOKO:2163/ctxtAUFPAR-PS_AUFART").Text = "ZPS1"
        Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").nodeContextMenu "000002"
        Session.findById("wnd[1]").sendVKey 0
    End If
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subIDENTIFICATION:SAPLCOKO:2816/txtCAUFVD-KTEXT").Text = "Mecânica 3"
    Session.findById("wnd[0]").sendVKey 0
    If Session.ActiveWindow.Text = "Project Builder: subprojeto " & FormularioDados.txtPEP.Value Then Session.findById("wnd[0]").sendVKey 0
    Session.findById("wnd[0]/tbar[0]/btn[3]").press
    If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
    Session.findById("wnd[0]/tbar[0]/btn[3]").press
    If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTRMN").Select
    If Session.ActiveWindow.Text = "Project Builder: subprojeto " & FormularioDados.txtPEP.Value Then
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTRMN/ssubSUBSCR_2100:SAPLCOKO:2110/ctxtCAUFVD-GSTRP").Text = ""
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTRMN/ssubSUBSCR_2100:SAPLCOKO:2110/ctxtCAUFVD-GLTRP").Text = ""
        Session.findById("wnd[0]").sendVKey 0
        If Session.findById("wnd[0]/sbar").Text = "Programação regressiva de prazos (entrar data de conclusão)" Then Session.findById("wnd[0]").sendVKey 12

        Do While Session.ActiveWindow.Text = "Aceitar RD"
        Session.findById("wnd[1]/usr/btnSPOP-OPTION2").press
        Loop
        Session.findById("wnd[0]/tbar[0]/btn[3]").press
        If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
        Session.findById("wnd[0]/tbar[0]/btn[3]").press
        If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
    End If
    If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
    
    If Session.ActiveWindow.Text = "Incluir diagrama de rede padrão" Then Session.findById("wnd[1]/usr/btnSPOP-VAROPTION2").press

        Do While Session.ActiveWindow.Text = "Aceitar RD"
        Session.findById("wnd[1]/usr/btnSPOP-OPTION2").press
        Loop
    Session.findById("wnd[0]/tbar[0]/btn[3]").press
    If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
    Session.findById("wnd[0]/tbar[0]/btn[3]").press
    If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
End If

'Inserir diagrama módulo 4
If FormularioDados.nrmodulos.Value <> "1 Módulo" And FormularioDados.nrmodulos.Value <> "2 Módulos" And FormularioDados.nrmodulos.Value <> "3 Módulos" Then
    Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
    Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").nodeContextMenu "000002"
    Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectContextMenuItem "COPY_NET"
    Session.findById("wnd[1]/usr/tabsTABSTR_2160/tabpTAB_TEMPLATE/ssubSUBSCREEN_2160:SAPLCOKO:2161/ctxtCAUFVD-STDNR").Text = Range("I8").Value
    Session.findById("wnd[1]/usr/tabsTABSTR_2160/tabpTAB_TEMPLATE/ssubSUBSCREEN_2160:SAPLCOKO:2161/txtCAUFVD-PLNAL").Text = Range("J8").Value
    Session.findById("wnd[1]").sendVKey 0
    If Session.findById("wnd[0]/sbar").Text = "Indicar um tipo de diagrama rede" Then
        Session.findById("wnd[1]/usr/tabsTABSTR_2160/tabpTAB_PARAMETER/ssubSUBSCREEN_2160:SAPLCOKO:2163/ctxtAUFPAR-PS_AUFART").Text = "ZPS1"
        Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").nodeContextMenu "000002"
        Session.findById("wnd[1]").sendVKey 0
    End If
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subIDENTIFICATION:SAPLCOKO:2816/txtCAUFVD-KTEXT").Text = "Mecânica 4"
    Session.findById("wnd[0]").sendVKey 0
    If Session.ActiveWindow.Text = "Project Builder: subprojeto " & FormularioDados.txtPEP.Value Then Session.findById("wnd[0]").sendVKey 0
    Session.findById("wnd[0]/tbar[0]/btn[3]").press
    If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
    Session.findById("wnd[0]/tbar[0]/btn[3]").press
    If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTRMN").Select
    If Session.ActiveWindow.Text = "Project Builder: subprojeto " & FormularioDados.txtPEP.Value Then
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTRMN/ssubSUBSCR_2100:SAPLCOKO:2110/ctxtCAUFVD-GSTRP").Text = ""
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTRMN/ssubSUBSCR_2100:SAPLCOKO:2110/ctxtCAUFVD-GLTRP").Text = ""
        Session.findById("wnd[0]").sendVKey 0
        If Session.findById("wnd[0]/sbar").Text = "Programação regressiva de prazos (entrar data de conclusão)" Then Session.findById("wnd[0]").sendVKey 12

        Do While Session.ActiveWindow.Text = "Aceitar RD"
        Session.findById("wnd[1]/usr/btnSPOP-OPTION2").press
        Loop
        Session.findById("wnd[0]/tbar[0]/btn[3]").press
        If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
        Session.findById("wnd[0]/tbar[0]/btn[3]").press
        If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
    End If
    If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
    
    If Session.ActiveWindow.Text = "Incluir diagrama de rede padrão" Then Session.findById("wnd[1]/usr/btnSPOP-VAROPTION2").press
    
        Do While Session.ActiveWindow.Text = "Aceitar RD"
        Session.findById("wnd[1]/usr/btnSPOP-OPTION2").press
        Loop
    Session.findById("wnd[0]/tbar[0]/btn[3]").press
    If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
    Session.findById("wnd[0]/tbar[0]/btn[3]").press
    If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
End If

'Inserir diagrama módulo 5
If FormularioDados.nrmodulos.Value <> "1 Módulo" And FormularioDados.nrmodulos.Value <> "2 Módulos" And FormularioDados.nrmodulos.Value <> "3 Módulos" And FormularioDados.nrmodulos.Value <> "4 Módulos" Then
    Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
    Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").nodeContextMenu "000002"
    Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectContextMenuItem "COPY_NET"
    Session.findById("wnd[1]/usr/tabsTABSTR_2160/tabpTAB_TEMPLATE/ssubSUBSCREEN_2160:SAPLCOKO:2161/ctxtCAUFVD-STDNR").Text = Range("I9").Value
    Session.findById("wnd[1]/usr/tabsTABSTR_2160/tabpTAB_TEMPLATE/ssubSUBSCREEN_2160:SAPLCOKO:2161/txtCAUFVD-PLNAL").Text = Range("J9").Value
    Session.findById("wnd[1]").sendVKey 0
    If Session.findById("wnd[0]/sbar").Text = "Indicar um tipo de diagrama rede" Then
        Session.findById("wnd[1]/usr/tabsTABSTR_2160/tabpTAB_PARAMETER/ssubSUBSCREEN_2160:SAPLCOKO:2163/ctxtAUFPAR-PS_AUFART").Text = "ZPS1"
        Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").nodeContextMenu "000002"
        Session.findById("wnd[1]").sendVKey 0
    End If
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subIDENTIFICATION:SAPLCOKO:2816/txtCAUFVD-KTEXT").Text = "Mecânica 5"
    Session.findById("wnd[0]").sendVKey 0
    If Session.ActiveWindow.Text = "Project Builder: subprojeto " & FormularioDados.txtPEP.Value Then Session.findById("wnd[0]").sendVKey 0
    Session.findById("wnd[0]/tbar[0]/btn[3]").press
    If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
    Session.findById("wnd[0]/tbar[0]/btn[3]").press
    If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTRMN").Select
    If Session.ActiveWindow.Text = "Project Builder: subprojeto " & FormularioDados.txtPEP.Value Then
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTRMN/ssubSUBSCR_2100:SAPLCOKO:2110/ctxtCAUFVD-GSTRP").Text = ""
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTRMN/ssubSUBSCR_2100:SAPLCOKO:2110/ctxtCAUFVD-GLTRP").Text = ""
        Session.findById("wnd[0]").sendVKey 0
        If Session.findById("wnd[0]/sbar").Text = "Programação regressiva de prazos (entrar data de conclusão)" Then Session.findById("wnd[0]").sendVKey 12

        Do While Session.ActiveWindow.Text = "Aceitar RD"
        Session.findById("wnd[1]/usr/btnSPOP-OPTION2").press
        Loop
        Session.findById("wnd[0]/tbar[0]/btn[3]").press
        If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
        Session.findById("wnd[0]/tbar[0]/btn[3]").press
        If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
    End If
    If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
    
    If Session.ActiveWindow.Text = "Incluir diagrama de rede padrão" Then Session.findById("wnd[1]/usr/btnSPOP-VAROPTION2").press
    
        Do While Session.ActiveWindow.Text = "Aceitar RD"
        Session.findById("wnd[1]/usr/btnSPOP-OPTION2").press
        Loop
    Session.findById("wnd[0]/tbar[0]/btn[3]").press
    If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
    Session.findById("wnd[0]/tbar[0]/btn[3]").press
    If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
End If

'Inserir diagrama módulo 6
If FormularioDados.nrmodulos.Value <> "1 Módulo" And FormularioDados.nrmodulos.Value <> "2 Módulos" And FormularioDados.nrmodulos.Value <> "3 Módulos" And FormularioDados.nrmodulos.Value <> "4 Módulos" And FormularioDados.nrmodulos.Value <> "5 Módulos" Then
    Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
    Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").nodeContextMenu "000002"
    Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectContextMenuItem "COPY_NET"
    Session.findById("wnd[1]/usr/tabsTABSTR_2160/tabpTAB_TEMPLATE/ssubSUBSCREEN_2160:SAPLCOKO:2161/ctxtCAUFVD-STDNR").Text = Range("I10").Value
    Session.findById("wnd[1]/usr/tabsTABSTR_2160/tabpTAB_TEMPLATE/ssubSUBSCREEN_2160:SAPLCOKO:2161/txtCAUFVD-PLNAL").Text = Range("J10").Value
    Session.findById("wnd[1]").sendVKey 0
    If Session.findById("wnd[0]/sbar").Text = "Indicar um tipo de diagrama rede" Then
        Session.findById("wnd[1]/usr/tabsTABSTR_2160/tabpTAB_PARAMETER/ssubSUBSCREEN_2160:SAPLCOKO:2163/ctxtAUFPAR-PS_AUFART").Text = "ZPS1"
        Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").nodeContextMenu "000002"
        Session.findById("wnd[1]").sendVKey 0
    End If
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subIDENTIFICATION:SAPLCOKO:2816/txtCAUFVD-KTEXT").Text = "Mecânica 6"
    Session.findById("wnd[0]").sendVKey 0
    If Session.ActiveWindow.Text = "Project Builder: subprojeto " & FormularioDados.txtPEP.Value Then Session.findById("wnd[0]").sendVKey 0
    Session.findById("wnd[0]/tbar[0]/btn[3]").press
    If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
    Session.findById("wnd[0]/tbar[0]/btn[3]").press
    If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTRMN").Select
    If Session.ActiveWindow.Text = "Project Builder: subprojeto " & FormularioDados.txtPEP.Value Then
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTRMN/ssubSUBSCR_2100:SAPLCOKO:2110/ctxtCAUFVD-GSTRP").Text = ""
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTRMN/ssubSUBSCR_2100:SAPLCOKO:2110/ctxtCAUFVD-GLTRP").Text = ""
        Session.findById("wnd[0]").sendVKey 0
        If Session.findById("wnd[0]/sbar").Text = "Programação regressiva de prazos (entrar data de conclusão)" Then Session.findById("wnd[0]").sendVKey 12

        Do While Session.ActiveWindow.Text = "Aceitar RD"
        Session.findById("wnd[1]/usr/btnSPOP-OPTION2").press
        Loop
        Session.findById("wnd[0]/tbar[0]/btn[3]").press
        If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
        Session.findById("wnd[0]/tbar[0]/btn[3]").press
        If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
    End If
    If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
    
    If Session.ActiveWindow.Text = "Incluir diagrama de rede padrão" Then Session.findById("wnd[1]/usr/btnSPOP-VAROPTION2").press
    
        Do While Session.ActiveWindow.Text = "Aceitar RD"
        Session.findById("wnd[1]/usr/btnSPOP-OPTION2").press
        Loop
    Session.findById("wnd[0]/tbar[0]/btn[3]").press
    If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
    Session.findById("wnd[0]/tbar[0]/btn[3]").press
    If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
End If

'Inserir diagrama módulo 7
If FormularioDados.nrmodulos.Value <> "1 Módulo" And FormularioDados.nrmodulos.Value <> "2 Módulos" And FormularioDados.nrmodulos.Value <> "3 Módulos" And FormularioDados.nrmodulos.Value <> "4 Módulos" And FormularioDados.nrmodulos.Value <> "5 Módulos" And FormularioDados.nrmodulos.Value <> "6 Módulos" Then
    Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
    Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").nodeContextMenu "000002"
    Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectContextMenuItem "COPY_NET"
    Session.findById("wnd[1]/usr/tabsTABSTR_2160/tabpTAB_TEMPLATE/ssubSUBSCREEN_2160:SAPLCOKO:2161/ctxtCAUFVD-STDNR").Text = Range("I11").Value
    Session.findById("wnd[1]/usr/tabsTABSTR_2160/tabpTAB_TEMPLATE/ssubSUBSCREEN_2160:SAPLCOKO:2161/txtCAUFVD-PLNAL").Text = Range("J11").Value
    Session.findById("wnd[1]").sendVKey 0
    If Session.findById("wnd[0]/sbar").Text = "Indicar um tipo de diagrama rede" Then
        Session.findById("wnd[1]/usr/tabsTABSTR_2160/tabpTAB_PARAMETER/ssubSUBSCREEN_2160:SAPLCOKO:2163/ctxtAUFPAR-PS_AUFART").Text = "ZPS1"
        Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").nodeContextMenu "000002"
        Session.findById("wnd[1]").sendVKey 0
    End If
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subIDENTIFICATION:SAPLCOKO:2816/txtCAUFVD-KTEXT").Text = "Mecânica 7"
    Session.findById("wnd[0]").sendVKey 0
    If Session.ActiveWindow.Text = "Project Builder: subprojeto " & FormularioDados.txtPEP.Value Then Session.findById("wnd[0]").sendVKey 0
    Session.findById("wnd[0]/tbar[0]/btn[3]").press
    If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
    Session.findById("wnd[0]/tbar[0]/btn[3]").press
    If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTRMN").Select
    If Session.ActiveWindow.Text = "Project Builder: subprojeto " & FormularioDados.txtPEP.Value Then
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTRMN/ssubSUBSCR_2100:SAPLCOKO:2110/ctxtCAUFVD-GSTRP").Text = ""
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTRMN/ssubSUBSCR_2100:SAPLCOKO:2110/ctxtCAUFVD-GLTRP").Text = ""
        Session.findById("wnd[0]").sendVKey 0
        If Session.findById("wnd[0]/sbar").Text = "Programação regressiva de prazos (entrar data de conclusão)" Then Session.findById("wnd[0]").sendVKey 12

        Do While Session.ActiveWindow.Text = "Aceitar RD"
        Session.findById("wnd[1]/usr/btnSPOP-OPTION2").press
        Loop
        Session.findById("wnd[0]/tbar[0]/btn[3]").press
        If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
        Session.findById("wnd[0]/tbar[0]/btn[3]").press
        If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
    End If
    If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
    
    If Session.ActiveWindow.Text = "Incluir diagrama de rede padrão" Then Session.findById("wnd[1]/usr/btnSPOP-VAROPTION2").press
    
        Do While Session.ActiveWindow.Text = "Aceitar RD"
        Session.findById("wnd[1]/usr/btnSPOP-OPTION2").press
        Loop
    Session.findById("wnd[0]/tbar[0]/btn[3]").press
    If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
    Session.findById("wnd[0]/tbar[0]/btn[3]").press
    If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
End If

'Inserir diagrama módulo 8
If FormularioDados.nrmodulos.Value <> "1 Módulo" And FormularioDados.nrmodulos.Value <> "2 Módulos" And FormularioDados.nrmodulos.Value <> "3 Módulos" And FormularioDados.nrmodulos.Value <> "4 Módulos" And FormularioDados.nrmodulos.Value <> "5 Módulos" And FormularioDados.nrmodulos.Value <> "6 Módulos" And FormularioDados.nrmodulos.Value <> "7 Módulos" Then
    Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
    Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").nodeContextMenu "000002"
    Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectContextMenuItem "COPY_NET"
    Session.findById("wnd[1]/usr/tabsTABSTR_2160/tabpTAB_TEMPLATE/ssubSUBSCREEN_2160:SAPLCOKO:2161/ctxtCAUFVD-STDNR").Text = Range("I12").Value
    Session.findById("wnd[1]/usr/tabsTABSTR_2160/tabpTAB_TEMPLATE/ssubSUBSCREEN_2160:SAPLCOKO:2161/txtCAUFVD-PLNAL").Text = Range("J12").Value
    Session.findById("wnd[1]").sendVKey 0
    If Session.findById("wnd[0]/sbar").Text = "Indicar um tipo de diagrama rede" Then
        Session.findById("wnd[1]/usr/tabsTABSTR_2160/tabpTAB_PARAMETER/ssubSUBSCREEN_2160:SAPLCOKO:2163/ctxtAUFPAR-PS_AUFART").Text = "ZPS1"
        Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").nodeContextMenu "000002"
        Session.findById("wnd[1]").sendVKey 0
    End If
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subIDENTIFICATION:SAPLCOKO:2816/txtCAUFVD-KTEXT").Text = "Mecânica 8"
    Session.findById("wnd[0]").sendVKey 0
    If Session.ActiveWindow.Text = "Project Builder: subprojeto " & FormularioDados.txtPEP.Value Then Session.findById("wnd[0]").sendVKey 0
    Session.findById("wnd[0]/tbar[0]/btn[3]").press
    If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
    Session.findById("wnd[0]/tbar[0]/btn[3]").press
    If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTRMN").Select
    If Session.ActiveWindow.Text = "Project Builder: subprojeto " & FormularioDados.txtPEP.Value Then
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTRMN/ssubSUBSCR_2100:SAPLCOKO:2110/ctxtCAUFVD-GSTRP").Text = ""
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTRMN/ssubSUBSCR_2100:SAPLCOKO:2110/ctxtCAUFVD-GLTRP").Text = ""
        Session.findById("wnd[0]").sendVKey 0
        If Session.findById("wnd[0]/sbar").Text = "Programação regressiva de prazos (entrar data de conclusão)" Then Session.findById("wnd[0]").sendVKey 12

        Do While Session.ActiveWindow.Text = "Aceitar RD"
        Session.findById("wnd[1]/usr/btnSPOP-OPTION2").press
        Loop
        Session.findById("wnd[0]/tbar[0]/btn[3]").press
        If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
        Session.findById("wnd[0]/tbar[0]/btn[3]").press
        If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
    End If
    If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
    
    If Session.ActiveWindow.Text = "Incluir diagrama de rede padrão" Then Session.findById("wnd[1]/usr/btnSPOP-VAROPTION2").press
    
        Do While Session.ActiveWindow.Text = "Aceitar RD"
        Session.findById("wnd[1]/usr/btnSPOP-OPTION2").press
        Loop
    Session.findById("wnd[0]/tbar[0]/btn[3]").press
    If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
    Session.findById("wnd[0]/tbar[0]/btn[3]").press
    If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
End If

If FormularioDados.tipoestrutura.Value <> "Skid (mecânica)" And FormularioDados.Betim1310.Value = True And FormularioDados.tipoestrutura.Value <> "Fixo" And FormularioDados.tipoestrutura.Value = "Skid (com elêtrica)" Then
    SKID1310 = True
    Else
    SKID1310 = False
End If



'Inserir diagrama acessórios
If FormularioDados.tipoestrutura.Value <> "ESSW (elétrica)" And SKID1310 <> True And FormularioDados.tipoestrutura.Value <> "Pilotis" Then
    If SKID1310 = False Then
        Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
        Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").nodeContextMenu "000002"
        Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectContextMenuItem "COPY_NET"
        Session.findById("wnd[1]/usr/tabsTABSTR_2160/tabpTAB_TEMPLATE/ssubSUBSCREEN_2160:SAPLCOKO:2161/ctxtCAUFVD-STDNR").Text = Range("I13").Value
        Session.findById("wnd[1]/usr/tabsTABSTR_2160/tabpTAB_TEMPLATE/ssubSUBSCREEN_2160:SAPLCOKO:2161/txtCAUFVD-PLNAL").Text = Range("J13").Value
        Session.findById("wnd[1]").sendVKey 0
        If Session.findById("wnd[0]/sbar").Text = "Indicar um tipo de diagrama rede" Then
            Session.findById("wnd[1]/usr/tabsTABSTR_2160/tabpTAB_PARAMETER/ssubSUBSCREEN_2160:SAPLCOKO:2163/ctxtAUFPAR-PS_AUFART").Text = "ZPS1"
            Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").nodeContextMenu "000002"
            Session.findById("wnd[1]").sendVKey 0
        End If
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subIDENTIFICATION:SAPLCOKO:2816/txtCAUFVD-KTEXT").Text = "Acessórios"
        Session.findById("wnd[0]").sendVKey 0
        If Session.ActiveWindow.Text = "Project Builder: subprojeto " & FormularioDados.txtPEP.Value Then Session.findById("wnd[0]").sendVKey 0
        Session.findById("wnd[0]/tbar[0]/btn[3]").press
        If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
        Session.findById("wnd[0]/tbar[0]/btn[3]").press
        If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTRMN").Select
        If Session.ActiveWindow.Text = "Project Builder: subprojeto " & FormularioDados.txtPEP.Value Then
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTRMN/ssubSUBSCR_2100:SAPLCOKO:2110/ctxtCAUFVD-GSTRP").Text = ""
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTRMN/ssubSUBSCR_2100:SAPLCOKO:2110/ctxtCAUFVD-GLTRP").Text = ""
            Session.findById("wnd[0]").sendVKey 0
            If Session.findById("wnd[0]/sbar").Text = "Programação regressiva de prazos (entrar data de conclusão)" Then Session.findById("wnd[0]").sendVKey 12
        
            Do While Session.ActiveWindow.Text = "Aceitar RD"
            Session.findById("wnd[1]/usr/btnSPOP-OPTION2").press
            Loop
            Session.findById("wnd[0]/tbar[0]/btn[3]").press
            If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
            Session.findById("wnd[0]/tbar[0]/btn[3]").press
            If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
        End If
        If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
        
        If Session.ActiveWindow.Text = "Incluir diagrama de rede padrão" Then Session.findById("wnd[1]/usr/btnSPOP-VAROPTION2").press
        
            Do While Session.ActiveWindow.Text = "Aceitar RD"
             Session.findById("wnd[1]/usr/btnSPOP-OPTION2").press
            Loop
        Session.findById("wnd[0]/tbar[0]/btn[3]").press
        If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
        Session.findById("wnd[0]/tbar[0]/btn[3]").press
        If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
    End If
SKID1310 = False
End If


'Inserir diagrama elétrica
Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").nodeContextMenu "000002"
Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectContextMenuItem "COPY_NET"
Session.findById("wnd[1]/usr/tabsTABSTR_2160/tabpTAB_TEMPLATE/ssubSUBSCREEN_2160:SAPLCOKO:2161/ctxtCAUFVD-STDNR").Text = Range("I14").Value
Session.findById("wnd[1]/usr/tabsTABSTR_2160/tabpTAB_TEMPLATE/ssubSUBSCREEN_2160:SAPLCOKO:2161/txtCAUFVD-PLNAL").Text = Range("J14").Value
Session.findById("wnd[1]").sendVKey 0
    If Session.findById("wnd[0]/sbar").Text = "Indicar um tipo de diagrama rede" Then
        Session.findById("wnd[1]/usr/tabsTABSTR_2160/tabpTAB_PARAMETER/ssubSUBSCREEN_2160:SAPLCOKO:2163/ctxtAUFPAR-PS_AUFART").Text = "ZPS1"
        Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").nodeContextMenu "000002"
        Session.findById("wnd[1]").sendVKey 0
    End If
Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subIDENTIFICATION:SAPLCOKO:2816/txtCAUFVD-KTEXT").Text = "Elétrica"
Session.findById("wnd[0]").sendVKey 0
If Session.ActiveWindow.Text = "Project Builder: subprojeto " & FormularioDados.txtPEP.Value Then Session.findById("wnd[0]").sendVKey 0
Session.findById("wnd[0]/tbar[0]/btn[3]").press
If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
Session.findById("wnd[0]/tbar[0]/btn[3]").press
If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTRMN").Select
If Session.ActiveWindow.Text = "Project Builder: subprojeto " & FormularioDados.txtPEP.Value Then
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTRMN/ssubSUBSCR_2100:SAPLCOKO:2110/ctxtCAUFVD-GSTRP").Text = ""
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTRMN/ssubSUBSCR_2100:SAPLCOKO:2110/ctxtCAUFVD-GLTRP").Text = ""
    Session.findById("wnd[0]").sendVKey 0
    If Session.findById("wnd[0]/sbar").Text = "Programação regressiva de prazos (entrar data de conclusão)" Then Session.findById("wnd[0]").sendVKey 12

        Do While Session.ActiveWindow.Text = "Aceitar RD"
        Session.findById("wnd[1]/usr/btnSPOP-OPTION2").press
        Loop
        
    Session.findById("wnd[0]/tbar[0]/btn[3]").press
    If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
    Session.findById("wnd[0]/tbar[0]/btn[3]").press
    If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
    Session.findById("wnd[0]/tbar[0]/btn[3]").press
End If
If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press

        Do While Session.ActiveWindow.Text = "Aceitar RD"
        Session.findById("wnd[1]/usr/btnSPOP-OPTION2").press
        Loop
Session.findById("wnd[0]/tbar[0]/btn[3]").press
If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
Session.findById("wnd[0]/tbar[0]/btn[3]").press
If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press

If FormularioDados.proBTI.Value = True Then
    Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
    Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").nodeContextMenu "000002"
    Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectContextMenuItem "COPY_NET"
    Session.findById("wnd[1]/usr/tabsTABSTR_2160/tabpTAB_TEMPLATE/ssubSUBSCREEN_2160:SAPLCOKO:2161/ctxtCAUFVD-STDNR").Text = "10001647"
    Session.findById("wnd[1]/usr/tabsTABSTR_2160/tabpTAB_TEMPLATE/ssubSUBSCREEN_2160:SAPLCOKO:2161/txtCAUFVD-PLNAL").Text = "1"
    Session.findById("wnd[1]/usr/tabsTABSTR_2160/tabpTAB_PARAMETER").Select
    Session.findById("wnd[1]/usr/tabsTABSTR_2160/tabpTAB_PARAMETER/ssubSUBSCREEN_2160:SAPLCOKO:2163/ctxtCAUFVD-WERKS").Text = "1313"
    Session.findById("wnd[1]").sendVKey 0
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subIDENTIFICATION:SAPLCOKO:2816/txtCAUFVD-KTEXT").Text = "Produção BTI"
    Session.findById("wnd[0]").sendVKey 0
    If Session.ActiveWindow.Text = "Project Builder: subprojeto " & FormularioDados.txtPEP.Value Then Session.findById("wnd[0]").sendVKey 0
    Session.findById("wnd[0]/tbar[0]/btn[3]").press
    If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
    Session.findById("wnd[0]/tbar[0]/btn[3]").press
    If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTRMN").Select
    If Session.ActiveWindow.Text = "Project Builder: subprojeto " & FormularioDados.txtPEP.Value Then
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTRMN/ssubSUBSCR_2100:SAPLCOKO:2110/ctxtCAUFVD-GSTRP").Text = ""
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTRMN/ssubSUBSCR_2100:SAPLCOKO:2110/ctxtCAUFVD-GLTRP").Text = ""
        Session.findById("wnd[0]").sendVKey 0
        If Session.findById("wnd[0]/sbar").Text = "Programação regressiva de prazos (entrar data de conclusão)" Then Session.findById("wnd[0]").sendVKey 12
            Do While Session.ActiveWindow.Text = "Aceitar RD"
            Session.findById("wnd[1]/usr/btnSPOP-OPTION2").press
            Loop
            Session.findById("wnd[0]/tbar[0]/btn[3]").press
            If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
            Session.findById("wnd[0]/tbar[0]/btn[3]").press
            If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
            Session.findById("wnd[0]/tbar[0]/btn[3]").press
        End If
    If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
    
            Do While Session.ActiveWindow.Text = "Aceitar RD"
            Session.findById("wnd[1]/usr/btnSPOP-OPTION2").press
            Loop
    Session.findById("wnd[0]/tbar[0]/btn[3]").press
    If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
    Session.findById("wnd[0]/tbar[0]/btn[3]").press
    If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
End If

'Adicionar diagrama de encerramento
Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").nodeContextMenu "000002"
Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectContextMenuItem "COPY_NET"
    If FormularioDados.Betim1310.Value = True Then
        Session.findById("wnd[1]/usr/tabsTABSTR_2160/tabpTAB_TEMPLATE/ssubSUBSCREEN_2160:SAPLCOKO:2161/ctxtCAUFVD-STDNR").Text = 10001732
        Session.findById("wnd[1]/usr/tabsTABSTR_2160/tabpTAB_TEMPLATE/ssubSUBSCREEN_2160:SAPLCOKO:2161/txtCAUFVD-PLNAL").Text = 1
        Else
        Session.findById("wnd[1]/usr/tabsTABSTR_2160/tabpTAB_TEMPLATE/ssubSUBSCREEN_2160:SAPLCOKO:2161/ctxtCAUFVD-STDNR").Text = 10001671
        Session.findById("wnd[1]/usr/tabsTABSTR_2160/tabpTAB_TEMPLATE/ssubSUBSCREEN_2160:SAPLCOKO:2161/txtCAUFVD-PLNAL").Text = 1
    End If
Session.findById("wnd[1]").sendVKey 0
    If Session.findById("wnd[0]/sbar").Text = "Indicar um tipo de diagrama rede" Then
        Session.findById("wnd[1]/usr/tabsTABSTR_2160/tabpTAB_PARAMETER/ssubSUBSCREEN_2160:SAPLCOKO:2163/ctxtAUFPAR-PS_AUFART").Text = "ZPS1"
        Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").nodeContextMenu "000002"
        Session.findById("wnd[1]").sendVKey 0
    End If
Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subIDENTIFICATION:SAPLCOKO:2816/txtCAUFVD-KTEXT").Text = "Encerramento"
Session.findById("wnd[0]").sendVKey 0
If Session.ActiveWindow.Text = "Project Builder: subprojeto " & FormularioDados.txtPEP.Value Then Session.findById("wnd[0]").sendVKey 0
Session.findById("wnd[0]/tbar[0]/btn[3]").press
If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
Session.findById("wnd[0]/tbar[0]/btn[3]").press
If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTRMN").Select
If Session.ActiveWindow.Text = "Project Builder: subprojeto " & FormularioDados.txtPEP.Value Then
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTRMN/ssubSUBSCR_2100:SAPLCOKO:2110/ctxtCAUFVD-GSTRP").Text = ""
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTRMN/ssubSUBSCR_2100:SAPLCOKO:2110/ctxtCAUFVD-GLTRP").Text = ""
    Session.findById("wnd[0]").sendVKey 0
    If Session.findById("wnd[0]/sbar").Text = "Programação regressiva de prazos (entrar data de conclusão)" Then Session.findById("wnd[0]").sendVKey 12

        Do While Session.ActiveWindow.Text = "Aceitar RD"
        Session.findById("wnd[1]/usr/btnSPOP-OPTION2").press
        Loop
        
    Session.findById("wnd[0]/tbar[0]/btn[3]").press
    If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
    Session.findById("wnd[0]/tbar[0]/btn[3]").press
    If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
    Session.findById("wnd[0]/tbar[0]/btn[3]").press
End If
If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press

        Do While Session.ActiveWindow.Text = "Aceitar RD"
        Session.findById("wnd[1]/usr/btnSPOP-OPTION2").press
        Loop
Session.findById("wnd[0]/tbar[0]/btn[3]").press
If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press
Session.findById("wnd[0]/tbar[0]/btn[3]").press
If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press




'Salvar
Session.findById("wnd[0]/tbar[0]/btn[11]").press

' Aviso agr de diagramas
If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Agrupamento de Diagramas" Then Session.findById("wnd[1]/usr/btnBUTTON_2").press
If Session.ActiveWindow.Text = "Agrupamento de Diagramas" Then Session.findById("wnd[1]/usr/btnBUTTON_2").press
If Session.ActiveWindow.Text = "Agrupamento de Diagramas" Then Session.findById("wnd[1]/usr/btnBUTTON_2").press
If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Agrupamento de Diagramas" Then Session.findById("wnd[1]/usr/btnBUTTON_2").press
If Session.ActiveWindow.Text = "Agrupamento de Diagramas" Then Session.findById("wnd[1]/usr/btnBUTTON_2").press
If Session.ActiveWindow.Text = "Agrupamento de Diagramas" Then Session.findById("wnd[1]/usr/btnBUTTON_2").press
If Session.ActiveWindow.Text = "Agrupamento de Diagramas" Then Session.findById("wnd[1]/usr/btnBUTTON_2").press
If Session.ActiveWindow.Text = "Agrupamento de Diagramas" Then Session.findById("wnd[1]/usr/btnBUTTON_2").press
If Session.ActiveWindow.Text = "Agrupamento de Diagramas" Then Session.findById("wnd[1]/usr/btnBUTTON_2").press
If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Agrupamento de Diagramas" Then Session.findById("wnd[1]/usr/btnBUTTON_2").press
If Session.ActiveWindow.Text = "Agrupamento de Diagramas" Then Session.findById("wnd[1]/usr/btnBUTTON_2").press
If Session.ActiveWindow.Text = "Agrupamento de Diagramas" Then Session.findById("wnd[1]/usr/btnBUTTON_2").press
If Session.ActiveWindow.Text = "Agrupamento de Diagramas" Then Session.findById("wnd[1]/usr/btnBUTTON_2").press

'Aviso orçamento de custos CPC47
If Session.ActiveWindow.Text = "Ctrl.disponibilidade" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press

'Erro na determinação dos custos
If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press


'Aviso Bloco K
If Session.ActiveWindow.Text = "Baixa existente em período anterior" Then Session.findById("wnd[1]/usr/btnBUTTON_1").press

'Sair da CJ20N
Session.findById("wnd[0]/tbar[0]/btn[15]").press


'Iniciar próximo procedimento
Call ajustar_tarefas

End Sub
```

#### `ajustar_tarefas` (Sub)

```vb
Sub ajustar_tarefas()

Dim SapGuiAuto As Object
Dim Application As Object
Dim Connection As Object
Dim Session As Object
Dim WScript As Object
Dim CT As String
Dim Nomenc As String
Dim j As Long
Dim sessao_ok As Boolean

'Conexão com o Objeto SAP
Set SapGuiAuto = GetObject("SAPGUI")
Set Application = SapGuiAuto.GetScriptingEngine
Set Connection = Application.Children(0)

sessao_ok = False
For j = 0 To Application.Children(0).Sessions.Count() - 1
    Set Session = Connection.Children(CLng(j))
    If Session.ActiveWindow.Text = "SAP Easy Access" Then
        sessao_ok = True
        Exit For
    End If
Next

If sessao_ok = False Then
    MsgBox "Nenhuma janela do SAP na tela inicial foi encontrada. Programa interrompido.", vbOKOnly
    Exit Sub
End If

ThisWorkbook.Application.StatusBar = "PEP " & FormularioDados.txtPEP.Value & ": Ajustando diagramas de rede"

'Abrir CJ20N e inserir PEP
Session.findById("wnd[0]").maximize
Session.findById("wnd[0]/tbar[0]/okcd").Text = "/NCJ20N"
Session.findById("wnd[0]").sendVKey 0
Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[0]/shell").pressButton "OPEN"
Session.findById("wnd[1]/usr/ctxtCNPB_W_ADD_OBJ_DYN-PROJ_EXT").Text = ""
Session.findById("wnd[1]/usr/ctxtCNPB_W_ADD_OBJ_DYN-PRPS_EXT").Text = FormularioDados.txtPEP.Value
Session.findById("wnd[1]/tbar[0]/btn[0]").press

'Por em modo de edição
If Session.ActiveWindow.Text = "Project Builder: exibir subprojeto " & FormularioDados.txtPEP.Value Then
    Session.findById("wnd[0]/tbar[1]/btn[13]").press
    If Session.findById("wnd[0]/sbar").Text = "Não foram bloqueados todos os objetos (ver protocolo de bloqueio)" Then
        Session.findById("wnd[0]/tbar[0]/btn[15]").press
        MsgBox "PEP está aberto por outro usuário. Verificar e tentar novamente mais tarde.", vbOKOnly
        Exit Sub
    End If
Else
    If Session.findById("wnd[0]/sbar").Text = "Não foram bloqueados todos os objetos (ver protocolo de bloqueio)" Then
        Session.findById("wnd[0]/tbar[0]/btn[15]").press
        MsgBox "PEP está aberto por outro usuário. Verificar e tentar novamente mais tarde.", vbOKOnly
        Exit Sub
    End If
End If

'Eliminar PRE e SES se for eletrocentro móvel
Dim pre_ok As Boolean
Dim ses_ok As Boolean
Dim nr_linhas_visiveis As Integer
Dim nr_max_linhas As Integer
Dim x As Integer
Dim y As Integer

pre_ok = False
ses_ok = False

If FormularioDados.tipoestrutura.Value = "Móvel" And FormularioDados.SemEng.Value = False Then
    Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "ACTY_OVW"
    nr_max_linhas = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Maximum + 1
    nr_linhas_visiveis = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").visiblerowcount
    For x = nr_linhas_visiveis - 1 To 0 Step -1
        If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-VORNR[0," & x & "]").Text = "0748" Then
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(x).Selected = True
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/btnBUTTON_DELETE").press
            Session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press
            ses_ok = True
        End If
        If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-VORNR[0," & x & "]").Text = "0746" Then
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(x).Selected = True
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/btnBUTTON_DELETE").press
            Session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press
            pre_ok = True
        End If
    Next
    'Seguir procurando na próxima página caso não tenha encontrado as tarefas na primeira
    If pre_ok = False Or ses_ok = False Then
        If pre_ok = False Then
            If (nr_max_linhas - (nr_linhas_visiveis - 1)) > (2 * nr_linhas_visiveis) Then
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = nr_linhas_visiveis * 2
                y = nr_linhas_visiveis * 3
            Else
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = nr_linhas_visiveis
                y = nr_linhas_visiveis * 2
            End If
            For x = 0 To nr_max_linhas - y
                If x <> 0 Then
                    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position + 1
                End If
                If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-VORNR[0,0]").Text = "0746" Then
                    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position).Selected = True
                    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/btnBUTTON_DELETE").press
                    Session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press
                    pre_ok = True
                    Exit For
                End If
            Next
        End If
        If ses_ok = False Then
            If (nr_max_linhas - (nr_linhas_visiveis - 1)) > (2 * nr_linhas_visiveis) Then
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = nr_linhas_visiveis * 2
                y = nr_linhas_visiveis * 3
            Else
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = nr_linhas_visiveis
                y = nr_linhas_visiveis * 2
            End If
            For x = 0 To nr_max_linhas - y
                If x <> 0 Then
                    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position + 1
                End If
                If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-VORNR[0,0]").Text = "0748" Then
                    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position).Selected = True
                    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/btnBUTTON_DELETE").press
                    Session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press
                    ses_ok = True
                    Exit For
                End If
            Next
        End If
    End If
End If

'Alterar Expedição e eliminar Faturamento se for item filho
Dim exp_ok As Boolean
Dim FAT_ok As Boolean

exp_ok = False
FAT_ok = False

Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "NETW_OVW"
nr_max_linhas_1 = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010").verticalScrollbar.Maximum + 1
    For f = 0 To nr_max_linhas_1 - 1
        If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & f & "]").Text = "Encerramento" Then
            DIAGENC = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtNETW_OVW-AUFNR[0," & f & "]").Text
            Exit For
        End If
    Next f


If FormularioDados.chkFilho.Value = True And FormularioDados.tipoestrutura.Value <> "ESSW (elétrica)" Then
    Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "ACTY_OVW"
    nr_max_linhas = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Maximum + 1
    nr_linhas_visiveis = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").visiblerowcount
    For x = 0 To nr_linhas_visiveis - 1
        If exp_ok = False Then
            If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-VORNR[0," & x & "]").Text = "0900" Then
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-VORNR[0," & x & "]").Text = "0899"
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(x).Selected = True
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/btnBUTTON_RELATION_OVERVIEW").press
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2403/tabsTABSTRIP_2401/tabpAOBG/ssubSUBSCR_2401:SAPLCOVG:2404/tblSAPLCOVGTCTRL_2402/chkAFABD-NCHKZ[2,1]").Selected = True
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2403/tabsTABSTRIP_2401/tabpAOBG/ssubSUBSCR_2401:SAPLCOVG:2404/tblSAPLCOVGTCTRL_2402/ctxtAFABD-NETZPLAN[1,1]").Text = DIAGENC
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2403/tabsTABSTRIP_2401/tabpAOBG/ssubSUBSCR_2401:SAPLCOVG:2404/tblSAPLCOVGTCTRL_2402/ctxtAFABD-VORNR[0,1]").Text = "920"
                Session.findById("wnd[0]").sendVKey 0
                exp_ok = True
                Exit For
            End If
        End If
    Next
    Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "ACTY_OVW"
    For x = 0 To nr_linhas_visiveis - 1
        If FAT_ok = False Then
            If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-VORNR[0," & x & "]").Text = "0910" Or Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-VORNR[0," & x & "]").Text = "0310" Then
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(x).Selected = True
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/btnBUTTON_DELETE").press
                Session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press
                FAT_ok = True
                Exit For
            End If
        End If
    Next
    'Seguir procurando nas próximas páginas caso não tenha encontrado as tarefas na primeira
    If exp_ok = False Or FAT_ok = False Then
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = nr_max_linhas - nr_linhas_visiveis - 3
        For x = 0 To 3
            If x <> 0 Then
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position + 1
            End If
            If exp_ok = False Then
                If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-VORNR[0,0]").Text = "0900" Then
                    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-VORNR[0,0]").Text = "0899"
                    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position).Selected = True
                    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/btnBUTTON_RELATION_OVERVIEW").press
                    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2403/tabsTABSTRIP_2401/tabpAOBG/ssubSUBSCR_2401:SAPLCOVG:2404/tblSAPLCOVGTCTRL_2402/ctxtAFABD-NETZPLAN[1,5]").Text = DIAGENC
                    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2403/tabsTABSTRIP_2401/tabpAOBG/ssubSUBSCR_2401:SAPLCOVG:2404/tblSAPLCOVGTCTRL_2402/chkAFABD-NCHKZ[2,5]").Selected = True
                    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2403/tabsTABSTRIP_2401/tabpAOBG/ssubSUBSCR_2401:SAPLCOVG:2404/tblSAPLCOVGTCTRL_2402/ctxtAFABD-VORNR[0,5]").Text = "920"
                    Session.findById("wnd[0]").sendVKey 0
                    exp_ok = True
                    Exit For
                End If
            End If
        Next
        Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "ACTY_OVW"
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = nr_max_linhas - nr_linhas_visiveis - 3
        For x = 0 To 3
            If x <> 0 Then
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position + 1
            End If
            If FAT_ok = False Then
                If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-VORNR[0,0]").Text = "0910" Or Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-VORNR[0,0]").Text = "0310" Then
                    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position).Selected = True
                    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/btnBUTTON_DELETE").press
                    Session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press
                    FAT_ok = True
                    Exit For
                End If
            End If
        Next
    End If
Else
    If FormularioDados.chkFilho.Value = True And FormularioDados.tipoestrutura.Value = "ESSW (elétrica)" Then
        Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "ACTY_OVW"
        nr_max_linhas = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Maximum + 1
        nr_linhas_visiveis = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").visiblerowcount
        For x = 0 To nr_linhas_visiveis - 1
            If exp_ok = False Then
                If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-VORNR[0," & x & "]").Text = "0300" Then
                    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-VORNR[0," & x & "]").Text = "0299"
                    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(x).Selected = True
                    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/btnBUTTON_RELATION_OVERVIEW").press
                    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2403/tabsTABSTRIP_2401/tabpAOBG/ssubSUBSCR_2401:SAPLCOVG:2404/tblSAPLCOVGTCTRL_2402/chkAFABD-NCHKZ[2,1]").Selected = True
                    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2403/tabsTABSTRIP_2401/tabpAOBG/ssubSUBSCR_2401:SAPLCOVG:2404/tblSAPLCOVGTCTRL_2402/ctxtAFABD-VORNR[0,1]").Text = "320"
                    Session.findById("wnd[0]").sendVKey 0
                    exp_ok = True
                    Exit For
                End If
            End If
        Next
        Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "ACTY_OVW"
        For x = 0 To nr_linhas_visiveis - 1
            If FAT_ok = False Then
                If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-VORNR[0," & x & "]").Text = "0310" Then
                    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(x).Selected = True
                    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/btnBUTTON_DELETE").press
                    Session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press
                    FAT_ok = True
                    Exit For
                End If
            End If
        Next
        'Seguir procurando nas próximas páginas caso não tenha encontrado as tarefas na primeira
        If exp_ok = False Or FAT_ok = False Then
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = nr_max_linhas - nr_linhas_visiveis - 3
            For x = 0 To 3
                If x <> 0 Then
                    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position + 1
                End If
                If exp_ok = False Then
                    If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-VORNR[0,0]").Text = "0300" Then
                        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-VORNR[0,0]").Text = "0299"
                        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position).Selected = True
                        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/btnBUTTON_RELATION_OVERVIEW").press
                        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2403/tabsTABSTRIP_2401/tabpAOBG/ssubSUBSCR_2401:SAPLCOVG:2404/tblSAPLCOVGTCTRL_2402/chkAFABD-NCHKZ[2,1]").Selected = True
                        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2403/tabsTABSTRIP_2401/tabpAOBG/ssubSUBSCR_2401:SAPLCOVG:2404/tblSAPLCOVGTCTRL_2402/ctxtAFABD-VORNR[0,1]").Text = "320"
                        Session.findById("wnd[0]").sendVKey 0
                        exp_ok = True
                        Exit For
                    End If
                End If
            Next
            Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "ACTY_OVW"
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = nr_max_linhas - nr_linhas_visiveis - 3
            For x = 0 To 3
                If x <> 0 Then
                    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position + 1
                End If
                If FAT_ok = False Then
                    If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-VORNR[0,0]").Text = "0310" Then
                        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position).Selected = True
                        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/btnBUTTON_DELETE").press
                        Session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press
                        FAT_ok = True
                        Exit For
                    End If
                End If
            Next
        End If
    End If
End If

'Entra na sub para fazer a troca dos centros de trabalho quando for solar
If FormularioDados.chkSolar.Value = True Then
    Call alterar_CTs
End If

'Corrigir nomes dos diagramas
'#1
'Detectar qual é o diagrama diagrama da Mecânica
Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "NETW_OVW"
nr_max_linhas_1 = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010").verticalScrollbar.Maximum + 1
For x = 0 To nr_max_linhas_1 - 1
    If FormularioDados.tipoestrutura.Value <> "ESSW (elétrica)" Then
        If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & x & "]").Text <> "Engenharia Mecânica" Then
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & x & "]").SetFocus
            Session.findById("wnd[0]").sendVKey 2
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTEXT").Select
            If FormularioDados.SemEng.Value = False Then
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTEXT/ssubSUBSCR_2100:SAPLCOKO:2114/cntlTEXTEDITORH/shell").Text = "Engenharia Mecânica"
            Else
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTEXT/ssubSUBSCR_2100:SAPLCOKO:2114/cntlTEXTEDITORH/shell").Text = "Mecânica"
            End If
        End If
    Else
        If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & x & "]").Text <> "Elétrica" Then
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & x & "]").SetFocus
            Session.findById("wnd[0]").sendVKey 2
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTEXT").Select
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTEXT/ssubSUBSCR_2100:SAPLCOKO:2114/cntlTEXTEDITORH/shell").Text = "Elétrica"
        End If
    End If
    Exit For
Next

'#2
Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "NETW_OVW"
nr_max_linhas_1 = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010").verticalScrollbar.Maximum + 1
For x = 1 To nr_max_linhas_1 - 1
    If FormularioDados.tipoestrutura.Value <> "ESSW (elétrica)" And FormularioDados.tipoestrutura.Value <> "Skid (mecânica)" And FormularioDados.tipoestrutura.Value <> "Pilotis" Then
        If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & x & "]").Text <> "Engenharia Elétrica" Then
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & x & "]").SetFocus
            Session.findById("wnd[0]").sendVKey 2
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTEXT").Select
            If FormularioDados.SemEng.Value = False Then
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTEXT/ssubSUBSCR_2100:SAPLCOKO:2114/cntlTEXTEDITORH/shell").Text = "Engenharia Elétrica"
                Else
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTEXT/ssubSUBSCR_2100:SAPLCOKO:2114/cntlTEXTEDITORH/shell").Text = "Acessórios"
            End If
        End If
    Else
    Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "NETW_OVW"
        If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & x & "]").Text <> "Mecânica" And FormularioDados.tipoestrutura.Value = "Skid (mecânica)" Then
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & x & "]").SetFocus
            Session.findById("wnd[0]").sendVKey 2
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTEXT").Select
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTEXT/ssubSUBSCR_2100:SAPLCOKO:2114/cntlTEXTEDITORH/shell").Text = "Mecânica"
            Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "NETW_OVW"
        End If
        If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & x & "]").Text <> "Elétrica" And FormularioDados.tipoestrutura.Value = "Pilotis" Then
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & x & "]").SetFocus
            Session.findById("wnd[0]").sendVKey 2
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTEXT").Select
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTEXT/ssubSUBSCR_2100:SAPLCOKO:2114/cntlTEXTEDITORH/shell").Text = "Elétrica"
         End If
    End If
    Exit For
Next

'#3
Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "NETW_OVW"
nr_max_linhas_1 = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010").verticalScrollbar.Maximum + 1
For x = 2 To nr_max_linhas_1 - 1
    If FormularioDados.tipoestrutura.Value <> "ESSW (elétrica)" And FormularioDados.tipoestrutura.Value <> "Skid (mecânica)" And FormularioDados.tipoestrutura.Value <> "Pilotis" Then
        If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & x & "]").Text <> "Mecânica 1" Then
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[[1," & x & "]").SetFocus
            Session.findById("wnd[0]").sendVKey 2
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTEXT").Select
             If FormularioDados.SemEng.Value = False Then
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTEXT/ssubSUBSCR_2100:SAPLCOKO:2114/cntlTEXTEDITORH/shell").Text = "Mecânica 1"
                Else
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTEXT/ssubSUBSCR_2100:SAPLCOKO:2114/cntlTEXTEDITORH/shell").Text = "Elétrica"
             End If
        End If
    Else
        If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & x & "]").Text <> "Elétrica" And (FormularioDados.tipoestrutura.Value = "Skid (mecânica)" And FormularioDados.Betim1310.Value = False) Then
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & x & "]").SetFocus
            Session.findById("wnd[0]").sendVKey 2
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTEXT").Select
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTEXT/ssubSUBSCR_2100:SAPLCOKO:2114/cntlTEXTEDITORH/shell").Text = "Elétrica"
        End If
    End If
    Exit For
Next

'#4
Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "NETW_OVW"
nr_max_linhas_1 = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010").verticalScrollbar.Maximum + 1
For x = 3 To nr_max_linhas_1 - 1
    If FormularioDados.nrmodulos.Value <> "1 Módulo" Then
        If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & x & "]").Text <> "Mecânica 2" Then
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & x & "]").SetFocus
            Session.findById("wnd[0]").sendVKey 2
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTEXT").Select
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTEXT/ssubSUBSCR_2100:SAPLCOKO:2114/cntlTEXTEDITORH/shell").Text = "Mecânica 2"
        End If
    Else
        If FormularioDados.tipoestrutura.Value <> "ESSW (mecânica)" And FormularioDados.tipoestrutura.Value <> "ESSW (elétrica)" And FormularioDados.tipoestrutura.Value <> "Skid (mecânica)" And FormularioDados.tipoestrutura.Value <> "Pilotis" And FormularioDados.SemEng.Value = False Then
            If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & x & "]").Text <> "Acessórios" Then
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & x & "]").SetFocus
                Session.findById("wnd[0]").sendVKey 2
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTEXT").Select
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTEXT/ssubSUBSCR_2100:SAPLCOKO:2114/cntlTEXTEDITORH/shell").Text = "Acessórios"
            End If
        End If
    End If
    Exit For
 Next
 
'#5
Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "NETW_OVW"
nr_max_linhas_1 = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010").verticalScrollbar.Maximum + 1
For x = 4 To nr_max_linhas_1 - 1
    If FormularioDados.nrmodulos.Value <> "1 Módulo" And FormularioDados.nrmodulos.Value <> "2 Módulos" Then
        If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & x & "]").Text <> "Mecânica 3" Then
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & x & "]").SetFocus
            Session.findById("wnd[0]").sendVKey 2
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTEXT").Select
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTEXT/ssubSUBSCR_2100:SAPLCOKO:2114/cntlTEXTEDITORH/shell").Text = "Mecânica 3"
        End If
    Else
        If FormularioDados.tipoestrutura.Value <> "ESSW (mecânica)" And FormularioDados.tipoestrutura.Value <> "ESSW (elétrica)" And FormularioDados.tipoestrutura.Value <> "Skid (mecânica)" And FormularioDados.tipoestrutura.Value <> "Pilotis" And FormularioDados.nrmodulos.Value <> "2 Módulos" And FormularioDados.SemEng.Value = False Then
            If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & x & "]").Text <> "Elétrica" Then
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & x & "]").SetFocus
                Session.findById("wnd[0]").sendVKey 2
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTEXT").Select
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTEXT/ssubSUBSCR_2100:SAPLCOKO:2114/cntlTEXTEDITORH/shell").Text = "Elétrica"
            End If
        End If
        If FormularioDados.nrmodulos.Value = "2 Módulos" Then
            If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & x & "]").Text <> "Acessórios" Then
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & x & "]").SetFocus
                Session.findById("wnd[0]").sendVKey 2
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTEXT").Select
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTEXT/ssubSUBSCR_2100:SAPLCOKO:2114/cntlTEXTEDITORH/shell").Text = "Acessórios"
            End If
        End If
    End If
    Exit For
Next

'#6 ((Com 4 Módulos)
Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "NETW_OVW"
nr_max_linhas_1 = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010").verticalScrollbar.Maximum + 1
For x = 5 To nr_max_linhas_1 - 1
    If FormularioDados.nrmodulos.Value <> "1 Módulo" And FormularioDados.nrmodulos.Value <> "2 Módulos" And FormularioDados.nrmodulos.Value <> "3 Módulos" Then
        If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & x & "]").Text <> "Mecânica 4" Then
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[[1," & x & "]").SetFocus
            Session.findById("wnd[0]").sendVKey 2
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTEXT").Select
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTEXT/ssubSUBSCR_2100:SAPLCOKO:2114/cntlTEXTEDITORH/shell").Text = "Mecânica 4"
        End If
    Else
        If FormularioDados.nrmodulos.Value = "2 Módulos" Then
            If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & x & "]").Text <> "Elétrica" Then
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & x & "]").SetFocus
                Session.findById("wnd[0]").sendVKey 2
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTEXT").Select
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTEXT/ssubSUBSCR_2100:SAPLCOKO:2114/cntlTEXTEDITORH/shell").Text = "Elétrica"
            End If
        End If
        If FormularioDados.nrmodulos.Value = "3 Módulos" Then
            If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & x & "]").Text <> "Acessórios" Then
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & x & "]").SetFocus
                Session.findById("wnd[0]").sendVKey 2
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTEXT").Select
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTEXT/ssubSUBSCR_2100:SAPLCOKO:2114/cntlTEXTEDITORH/shell").Text = "Acessórios"
            End If
        End If
    End If
    Exit For
Next

'#7 (Com 5 Módulos)
Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "NETW_OVW"
nr_max_linhas_1 = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010").verticalScrollbar.Maximum + 1
For x = 6 To nr_max_linhas_1 - 1
    If FormularioDados.nrmodulos.Value <> "1 Módulo" And FormularioDados.nrmodulos.Value <> "2 Módulos" And FormularioDados.nrmodulos.Value <> "3 Módulos" And FormularioDados.nrmodulos.Value <> "4 Módulos" Then
        If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & x & "]").Text <> "Mecânica 5" Then
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & x & "]").SetFocus
            Session.findById("wnd[0]").sendVKey 2
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTEXT").Select
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTEXT/ssubSUBSCR_2100:SAPLCOKO:2114/cntlTEXTEDITORH/shell").Text = "Mecânica 5"
        End If
    Else
        If FormularioDados.nrmodulos.Value = "3 Módulos" Then
            If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & x & "]").Text <> "Elétrica" Then
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & x & "]").SetFocus
                Session.findById("wnd[0]").sendVKey 2
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTEXT").Select
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTEXT/ssubSUBSCR_2100:SAPLCOKO:2114/cntlTEXTEDITORH/shell").Text = "Elétrica"
            End If
        End If
        If FormularioDados.nrmodulos.Value = "4 Módulos" Then
            If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & x & "]").Text <> "Acessórios" Then
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & x & "]").SetFocus
                Session.findById("wnd[0]").sendVKey 2
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTEXT").Select
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTEXT/ssubSUBSCR_2100:SAPLCOKO:2114/cntlTEXTEDITORH/shell").Text = "Acessórios"
            End If
        End If
    End If
    Exit For
Next

'#8 (Com 6 Módulos)
Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "NETW_OVW"
nr_max_linhas_1 = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010").verticalScrollbar.Maximum + 1
For x = 7 To nr_max_linhas_1 - 1
    If FormularioDados.nrmodulos.Value <> "1 Módulo" And FormularioDados.nrmodulos.Value <> "2 Módulos" And FormularioDados.nrmodulos.Value <> "3 Módulos" And FormularioDados.nrmodulos.Value <> "4 Módulos" And FormularioDados.nrmodulos.Value <> "5 Módulos" Then
        If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & x & "]").Text <> "Mecânica 6" Then
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & x & "]").SetFocus
            Session.findById("wnd[0]").sendVKey 2
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTEXT").Select
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTEXT/ssubSUBSCR_2100:SAPLCOKO:2114/cntlTEXTEDITORH/shell").Text = "Mecânica 6"
        End If
    Else
        If FormularioDados.nrmodulos.Value = "4 Módulos" Then
            If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & x & "]").Text <> "Elétrica" Then
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & x & "]").SetFocus
                Session.findById("wnd[0]").sendVKey 2
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTEXT").Select
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTEXT/ssubSUBSCR_2100:SAPLCOKO:2114/cntlTEXTEDITORH/shell").Text = "Elétrica"
            End If
        End If
        If FormularioDados.nrmodulos.Value = "5 Módulos" Then
            If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & x & "]").Text <> "Acessórios" Then
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & x & "]").SetFocus
                Session.findById("wnd[0]").sendVKey 2
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTEXT").Select
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTEXT/ssubSUBSCR_2100:SAPLCOKO:2114/cntlTEXTEDITORH/shell").Text = "Acessórios"
            End If
        End If
    End If
    Exit For
Next

'#9 (Com 7 Módulos)
Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "NETW_OVW"
nr_max_linhas_1 = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010").verticalScrollbar.Maximum + 1
For x = 8 To nr_max_linhas_1 - 1
    If FormularioDados.nrmodulos.Value <> "1 Módulo" And FormularioDados.nrmodulos.Value <> "2 Módulos" And FormularioDados.nrmodulos.Value <> "3 Módulos" And FormularioDados.nrmodulos.Value <> "4 Módulos" And FormularioDados.nrmodulos.Value <> "5 Módulos" And FormularioDados.nrmodulos.Value <> "6 Módulos" Then
        If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & x & "]").Text <> "Mecânica 7" Then
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & x & "]").SetFocus
            Session.findById("wnd[0]").sendVKey 2
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTEXT").Select
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTEXT/ssubSUBSCR_2100:SAPLCOKO:2114/cntlTEXTEDITORH/shell").Text = "Mecânica 7"
        End If
    Else
        If FormularioDados.nrmodulos.Value = "5 Módulos" Then
            If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & x & "]").Text <> "Elétrica" Then
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & x & "]").SetFocus
                Session.findById("wnd[0]").sendVKey 2
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTEXT").Select
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTEXT/ssubSUBSCR_2100:SAPLCOKO:2114/cntlTEXTEDITORH/shell").Text = "Elétrica"
            End If
        End If
        If FormularioDados.nrmodulos.Value = "6 Módulos" Then
            If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & x & "]").Text <> "Acessórios" Then
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & x & "]").SetFocus
                Session.findById("wnd[0]").sendVKey 2
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTEXT").Select
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTEXT/ssubSUBSCR_2100:SAPLCOKO:2114/cntlTEXTEDITORH/shell").Text = "Acessórios"
            End If
        End If
    End If
    Exit For
Next

'#10 (Com 8 Módulos)
Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "NETW_OVW"
nr_max_linhas_1 = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010").verticalScrollbar.Maximum + 1
For x = 9 To nr_max_linhas_1 - 1
    Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "NETW_OVW"
    If FormularioDados.nrmodulos.Value <> "1 Módulo" And FormularioDados.nrmodulos.Value <> "2 Módulos" And FormularioDados.nrmodulos.Value <> "3 Módulos" And FormularioDados.nrmodulos.Value <> "4 Módulos" And FormularioDados.nrmodulos.Value <> "5 Módulos" And FormularioDados.nrmodulos.Value <> "6 Módulos" And FormularioDados.nrmodulos.Value <> "7 Módulos" Then
        If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & x & "]").Text <> "Mecânica 8" Then
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & x & "]").SetFocus
            Session.findById("wnd[0]").sendVKey 2
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTEXT").Select
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTEXT/ssubSUBSCR_2100:SAPLCOKO:2114/cntlTEXTEDITORH/shell").Text = "Mecânica 8"
        End If
    Else
        If FormularioDados.nrmodulos.Value = "6 Módulos" Then
            If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & x & "]").Text <> "Elétrica" Then
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & x & "]").SetFocus
                Session.findById("wnd[0]").sendVKey 2
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTEXT").Select
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTEXT/ssubSUBSCR_2100:SAPLCOKO:2114/cntlTEXTEDITORH/shell").Text = "Elétrica"
            End If
        End If
        If FormularioDados.nrmodulos.Value = "7 Módulos" Then
            If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & x & "]").Text <> "Acessórios" Then
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & x & "]").SetFocus
                Session.findById("wnd[0]").sendVKey 2
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTEXT").Select
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTEXT/ssubSUBSCR_2100:SAPLCOKO:2114/cntlTEXTEDITORH/shell").Text = "Acessórios"
            End If
        End If
    End If
    Exit For
Next

'#11
Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "NETW_OVW"
nr_max_linhas_1 = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010").verticalScrollbar.Maximum + 1
For x = 10 To nr_max_linhas_1 - 1
    Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "NETW_OVW"
    If FormularioDados.nrmodulos.Value = "4 Módulos" Then
        If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & x & "]").Text <> "Acessórios" Then
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & x & "]").SetFocus
            Session.findById("wnd[0]").sendVKey 2
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTEXT").Select
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTEXT/ssubSUBSCR_2100:SAPLCOKO:2114/cntlTEXTEDITORH/shell").Text = "Acessórios"
        End If
    End If
    If FormularioDados.nrmodulos.Value = "3 Módulos" Then
        If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & x & "]").Text <> "Elétrica" Then
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & x & "]").SetFocus
            Session.findById("wnd[0]").sendVKey 2
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTEXT").Select
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTEXT/ssubSUBSCR_2100:SAPLCOKO:2114/cntlTEXTEDITORH/shell").Text = "Elétrica"
        End If
    End If
    Exit For
Next

'#12
Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "NETW_OVW"
nr_max_linhas_1 = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010").verticalScrollbar.Maximum + 1
For x = 11 To nr_max_linhas_1 - 1
    If FormularioDados.nrmodulos.Value = "4 Módulos" Then
        If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & x & "]").Text <> "Elétrica" Then
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtCAUFVD-KTEXT[1," & x & "]").SetFocus
            Session.findById("wnd[0]").sendVKey 2
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTEXT").Select
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTEXT/ssubSUBSCR_2100:SAPLCOKO:2114/cntlTEXTEDITORH/shell").Text = "Elétrica"
        End If
    End If
    Exit For
Next
'Salvar
Session.findById("wnd[0]/tbar[0]/btn[11]").press

' Aviso agr de diagramas
If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Agrupamento de Diagramas" Then Session.findById("wnd[1]/usr/btnBUTTON_2").press
If Session.ActiveWindow.Text = "Agrupamento de Diagramas" Then Session.findById("wnd[1]/usr/btnBUTTON_2").press
If Session.ActiveWindow.Text = "Agrupamento de Diagramas" Then Session.findById("wnd[1]/usr/btnBUTTON_2").press
If Session.ActiveWindow.Text = "Agrupamento de Diagramas" Then Session.findById("wnd[1]/usr/btnBUTTON_2").press
If Session.ActiveWindow.Text = "Agrupamento de Diagramas" Then Session.findById("wnd[1]/usr/btnBUTTON_2").press
If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Agrupamento de Diagramas" Then Session.findById("wnd[1]/usr/btnBUTTON_2").press
If Session.ActiveWindow.Text = "Agrupamento de Diagramas" Then Session.findById("wnd[1]/usr/btnBUTTON_2").press
If Session.ActiveWindow.Text = "Agrupamento de Diagramas" Then Session.findById("wnd[1]/usr/btnBUTTON_2").press
If Session.ActiveWindow.Text = "Agrupamento de Diagramas" Then Session.findById("wnd[1]/usr/btnBUTTON_2").press
If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Agrupamento de Diagramas" Then Session.findById("wnd[1]/usr/btnBUTTON_2").press
If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Agrupamento de Diagramas" Then Session.findById("wnd[1]/usr/btnBUTTON_2").press
If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Agrupamento de Diagramas" Then Session.findById("wnd[1]/usr/btnBUTTON_2").press
If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Agrupamento de Diagramas" Then Session.findById("wnd[1]/usr/btnBUTTON_2").press
If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press

'Aviso orçamento de custos CPC47
If Session.ActiveWindow.Text = "Ctrl.disponibilidade" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press

If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press


'Erro na determinação dos custos
If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press


'Aviso Bloco K
If Session.ActiveWindow.Text = "Baixa existente em período anterior" Then Session.findById("wnd[1]/usr/btnBUTTON_1").press

'Sair da CJ20N
Session.findById("wnd[0]/tbar[0]/btn[15]").press

erro_DR = False

End Sub
```

#### `planejar` (Sub)

```vb
Sub planejar()

Dim SapGuiAuto As Object
Dim Application As Object
Dim Connection As Object
Dim Session As Object
Dim WScript As Object
Dim nr_linhas As Integer
Dim j As Long
Dim sessao_ok As Boolean

'Indentificar número de linhas de dados
ActiveSheet.Unprotect
Range("A3").Select
Selection.End(xlDown).Select
nr_linhas = Selection.Row - 3
ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True

'Conexão com o Objeto SAP
Set SapGuiAuto = GetObject("SAPGUI")
Set Application = SapGuiAuto.GetScriptingEngine
Set Connection = Application.Children(0)

sessao_ok = False
For j = 0 To Application.Children(0).Sessions.Count() - 1
    Set Session = Connection.Children(CLng(j))
    If Session.ActiveWindow.Text = "SAP Easy Access" Then
        sessao_ok = True
        Exit For
    End If
Next

If sessao_ok = False Then
    MsgBox "Nenhuma janela do SAP na tela inicial foi encontrada. Programa interrompido.", vbOKOnly
    Exit Sub
End If

ThisWorkbook.Application.StatusBar = "PEP " & FormularioDados.txtPEP.Value & ": Efetuando planejamento"


' Deletar tarefas de climatização se necessário
 'Sem climatização
'    If FormularioDados.tipomaq.Value = "Não possui" And FormularioDados.tipoestrutura.Value <> "ESSW (mecânica)" Then
'       Session.findById("wnd[0]/tbar[0]/okcd").Text = "/nCN47N"
'       Session.findById("wnd[0]").sendVKey 0
'        'Preencher perfil caso necessário
'            If Session.ActiveWindow.Text = "Entrar perfil" Then
'                Session.findById("wnd[1]/usr/ctxtTCNT-PROF_DB").Text = "000000000001"
'               Session.findById("wnd[1]/tbar[0]/btn[0]").press
'                    End If
'                        Session.findById("wnd[0]/usr/ctxtCN_PSPNR-LOW").Text = FormularioDados.txtPEP.Value'
'                        Session.findById("wnd[0]/usr/ctxtP_DISVAR").Text = ""
'                        Session.findById("wnd[0]/usr/ctxtCN_NETNR-LOW").Text = ""
'                        Session.findById("wnd[0]/usr/btn%_CN_ACTVT_%_APP_%-VALU_PUSH").press
'                        Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "0828"
'                        Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = "0830"
'                        Session.findById("wnd[1]/tbar[0]/btn[8]").press
'                        Session.findById("wnd[0]/usr/ctxtP_DISVAR").Text = "1SAP"
'                        Session.findById("wnd[0]/tbar[1]/btn[8]").press
'                        Session.findById("wnd[0]/usr/cntlALVCONTAINER/shellcont/shell").doubleClickCurrentCell
'                        Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell").topNode = "         23"
'                        Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").topNode = "000001"
'                        Session.findById("wnd[0]/tbar[1]/btn[13]").press
'                        Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[0]/shell").pressButton "DELE"
'                        Session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press
'                        Session.findById("wnd[0]/tbar[0]/btn[3]").press
'
'                        'Aviso orçamento de custos CPC47
'                        If Session.ActiveWindow.Text = "Ctrl.disponibilidade" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press''
'
'                        'Erro na determinação dos custos
'                        If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
'                        If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
'                        If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
'                        If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
'                        If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
'                        If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
'                        If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
'                        If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
'                        If Session.ActiveWindow.Text = "Ctrl.disponibilidade" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
'
'                        'Aviso Bloco K
'                        If Session.ActiveWindow.Text = "Baixa existente em período anterior" Then Session.findById("wnd[1]/usr/btnBUTTON_1").press
'
'                        Session.findById("wnd[0]/usr/cntlALVCONTAINER/shellcont/shell").currentCellRow = 1
'                        Session.findById("wnd[0]/usr/cntlALVCONTAINER/shellcont/shell").doubleClickCurrentCell
'                        Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell").topNode = "         23"
'                        Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").topNode = "000001"
'                        Session.findById("wnd[0]/tbar[1]/btn[13]").press
'                        Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[0]/shell").pressButton "DELE"
'                        Session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press
'                        Session.findById("wnd[0]/tbar[0]/btn[11]").press
'
'                        'Aviso orçamento de custos CPC47
'                        If Session.ActiveWindow.Text = "Ctrl.disponibilidade" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press'''
'
'                        'Erro na determinação dos custos
'                        If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
'                        If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
'                        If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
'                        If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
'                        If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
'                        If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
'                        If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
'                        If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
'                        If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
'
'                        'Aviso Bloco K
'                        If Session.ActiveWindow.Text = "Baixa existente em período anterior" Then Session.findById("wnd[1]/usr/btnBUTTON_1").press
'
'                        Session.findById("wnd[0]/tbar[0]/btn[3]").press
'                        Session.findById("wnd[0]/tbar[0]/btn[3]").press
'                End If
                            
 ' Deletar tarefa de Reles se não for necessário
'If FormularioDados.programacaoreles.Value = False And FormularioDados.tipoestrutura.Value = "Container Solar" Then
''Or FormularioDados.tipoestrutura.Value = "Fixo"
'Session.findById("wnd[0]/tbar[0]/okcd").Text = "/nCN47N"
'       Session.findById("wnd[0]").sendVKey 0
'        'Preencher perfil caso necessário
'            If Session.ActiveWindow.Text = "Entrar perfil" Then
'                Session.findById("wnd[1]/usr/ctxtTCNT-PROF_DB").Text = "000000000001"
'                Session.findById("wnd[1]/tbar[0]/btn[0]").press
'                    End If
'                        Session.findById("wnd[0]/usr/ctxtCN_PSPNR-LOW").Text = FormularioDados.txtPEP.Value
'                        Session.findById("wnd[0]/usr/ctxtP_DISVAR").Text = ""
'                        Session.findById("wnd[0]/usr/ctxtCN_NETNR-LOW").Text = ""
'                        Session.findById("wnd[0]/usr/btn%_CN_ACTVT_%_APP_%-VALU_PUSH").press
'                       Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "886"
'                        Session.findById("wnd[1]/tbar[0]/btn[8]").press
'                        Session.findById("wnd[0]/usr/ctxtP_DISVAR").Text = "1SAP"
'                        Session.findById("wnd[0]/tbar[1]/btn[8]").press
'                        Session.findById("wnd[0]/usr/cntlALVCONTAINER/shellcont/shell").doubleClickCurrentCell
'                        Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell").topNode = "         23"
'                        Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").topNode = "000001"
'                        Session.findById("wnd[0]/tbar[1]/btn[13]").press
'                        Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[0]/shell").pressButton "DELE"
 '                       Session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press
 '                       Session.findById("wnd[0]/tbar[0]/btn[3]").press
 '
 ''                       'Aviso orçamento de custos CPC47
 '                       If Session.ActiveWindow.Text = "Ctrl.disponibilidade" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press''
'
'                        'Erro na determinação dos custos
'                        If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
'                        If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
'                        If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
'                        If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
'                        If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
'                        If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
 '                       If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
'                        If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
 '                       If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
'
'                        'Aviso Bloco K
'                        If Session.ActiveWindow.Text = "Baixa existente em período anterior" Then Session.findById("wnd[1]/usr/btnBUTTON_1").press
'
'                        'Aviso orçamento de custos CPC47
 '                       If Session.ActiveWindow.Text = "Ctrl.disponibilidade" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
''
'                        'Erro na determinação dos custos
'                        If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
'                        If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
 ''                       If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
 '                       If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
 '                       If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
 '                       If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
  '                      If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
  '                      If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
  ''                      If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
  '
  '                      'Aviso Bloco K
  '                      If Session.ActiveWindow.Text = "Baixa existente em período anterior" Then Session.findById("wnd[1]/usr/btnBUTTON_1").press
  '
  '                      Session.findById("wnd[0]/tbar[0]/btn[3]").press
  '                      Session.findById("wnd[0]/tbar[0]/btn[3]").press

'End If

' Deletar tarefa de Edificação
 'Edf
 '   If FormularioDados.tipoestrutura.Value = "Container Solar" And FormularioDados.Betim1310.Value = False Then
 '      Session.findById("wnd[0]/tbar[0]/okcd").Text = "/nCN47N"
 '      Session.findById("wnd[0]").sendVKey 0
        'Preencher perfil caso necessário
 '           If Session.ActiveWindow.Text = "Entrar perfil" Then
 '               Session.findById("wnd[1]/usr/ctxtTCNT-PROF_DB").Text = "000000000001"
 '               Session.findById("wnd[1]/tbar[0]/btn[0]").press
 '               End If
 '                       Session.findById("wnd[0]/usr/ctxtCN_PSPNR-LOW").Text = FormularioDados.txtPEP.Value
 '                       Session.findById("wnd[0]/usr/ctxtP_DISVAR").Text = "/WAU-PCP_K"
 '                       Session.findById("wnd[0]/usr/ctxtCN_NETNR-LOW").Text = ""
 '                       Session.findById("wnd[0]/usr/btn%_CN_ACTVT_%_APP_%-VALU_PUSH").press
 '                       Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "750"
 '                       Session.findById("wnd[1]/tbar[0]/btn[8]").press
 '                       Session.findById("wnd[0]/usr/ctxtP_DISVAR").Text = "1SAP"
 '                       Session.findById("wnd[0]/tbar[1]/btn[8]").press
 '                       Session.findById("wnd[0]/usr/cntlALVCONTAINER/shellcont/shell").doubleClickCurrentCell
 '                       Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell").topNode = "         23"
 '                       Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").topNode = "000001"
 '                       Session.findById("wnd[0]/tbar[1]/btn[13]").press
 '                       Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[0]/shell").pressButton "DELE"
                        
                          
 '                       Session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press
 '                       Session.findById("wnd[0]/tbar[0]/btn[3]").press
                        
                        'Aviso orçamento de custos CPC47
 '                       If Session.ActiveWindow.Text = "Ctrl.disponibilidade" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press

                        'Erro na determinação dos custos
 '                       If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
 '                       If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
 '                       If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
 '                       If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
 '                       If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
 '                       If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
 '                       If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
 '                       If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
 '                       If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
                        
                        'Aviso Bloco K
 '                       If Session.ActiveWindow.Text = "Baixa existente em período anterior" Then Session.findById("wnd[1]/usr/btnBUTTON_1").press
                        
                        'Voltar a tela inicial
 '                       Session.findById("wnd[0]/tbar[0]/btn[3]").press
 '                       Session.findById("wnd[0]/tbar[0]/btn[3]").press
                        
'End If
                        
' Deletar fat ence
    If FormularioDados.proBTI.Value = True And FormularioDados.tipoestrutura.Value = "Container Solar" Then
       Session.findById("wnd[0]/tbar[0]/okcd").Text = "/nCN47N"
       Session.findById("wnd[0]").sendVKey 0
        'Preencher perfil caso necessário
            If Session.ActiveWindow.Text = "Entrar perfil" Then
                Session.findById("wnd[1]/usr/ctxtTCNT-PROF_DB").Text = "000000000001"
                Session.findById("wnd[1]/tbar[0]/btn[0]").press
                    End If
                        Session.findById("wnd[0]/usr/ctxtCN_PSPNR-LOW").Text = FormularioDados.txtPEP.Value
                        Session.findById("wnd[0]/usr/ctxtP_DISVAR").Text = ""
                        Session.findById("wnd[0]/usr/ctxtCN_NETNR-LOW").Text = ""
                        Session.findById("wnd[0]/usr/btn%_CN_ACTVT_%_APP_%-VALU_PUSH").press
                        Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "910"
                        Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = "920"
                        Session.findById("wnd[1]/tbar[0]/btn[8]").press
                        Session.findById("wnd[0]/usr/ctxtP_DISVAR").Text = "1SAP"
                        Session.findById("wnd[0]/tbar[1]/btn[8]").press
                        Session.findById("wnd[0]/usr/cntlALVCONTAINER/shellcont/shell").doubleClickCurrentCell
                        Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell").topNode = "         23"
                        Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").topNode = "000001"
                        Session.findById("wnd[0]/tbar[1]/btn[13]").press
                        Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[0]/shell").pressButton "DELE"
                        Session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press
                        Session.findById("wnd[0]/tbar[0]/btn[3]").press
                        
                        'Aviso orçamento de custos CPC47
                        If Session.ActiveWindow.Text = "Ctrl.disponibilidade" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press

                        'Erro na determinação dos custos
                        If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
                        If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
                        If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
                        If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
                        If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
                        If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
                        If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
                        If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
                        If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
                        
                        'Aviso Bloco K
                        If Session.ActiveWindow.Text = "Baixa existente em período anterior" Then Session.findById("wnd[1]/usr/btnBUTTON_1").press
               
                        Session.findById("wnd[0]/usr/cntlALVCONTAINER/shellcont/shell").currentCellRow = 1
                        Session.findById("wnd[0]/usr/cntlALVCONTAINER/shellcont/shell").doubleClickCurrentCell
                        Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell").topNode = "         23"
                        Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").topNode = "000001"
                        Session.findById("wnd[0]/tbar[1]/btn[13]").press
                        Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[0]/shell").pressButton "DELE"
                        Session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press
                        Session.findById("wnd[0]/tbar[0]/btn[11]").press
                        
                        'Aviso orçamento de custos CPC47
                        If Session.ActiveWindow.Text = "Ctrl.disponibilidade" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press

                        'Erro na determinação dos custos
                        If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
                        If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
                        If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
                        If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
                        If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
                        If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
                        If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
                        If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
                        If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
                        
                        'Aviso Bloco K
                        If Session.ActiveWindow.Text = "Baixa existente em período anterior" Then Session.findById("wnd[1]/usr/btnBUTTON_1").press
                                               
                        Session.findById("wnd[0]/tbar[0]/btn[3]").press
                        Session.findById("wnd[0]/tbar[0]/btn[3]").press
End If

'Abrir CJ20N e inserir PEP
Session.findById("wnd[0]").maximize
Session.findById("wnd[0]/tbar[0]/okcd").Text = "/NCJ20N"
Session.findById("wnd[0]").sendVKey 0
Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[0]/shell").pressButton "OPEN"
Session.findById("wnd[1]/usr/ctxtCNPB_W_ADD_OBJ_DYN-PROJ_EXT").Text = ""
Session.findById("wnd[1]/usr/ctxtCNPB_W_ADD_OBJ_DYN-PRPS_EXT").Text = FormularioDados.txtPEP.Value
Session.findById("wnd[1]").sendVKey 0

'Encerrar programa em caso de mensagem de erro (PEP inexistente)
If Session.ActiveWindow.Text = "Erro" Then
    Session.findById("wnd[2]").sendVKey 0
    Session.findById("wnd[1]").Close
    Session.findById("wnd[0]/tbar[0]/btn[15]").press
    MsgBox "PEP informado não existe. Programa interrompido.", vbOKOnly
        Exit Sub
    End If
'Por em modo de edição
If Session.ActiveWindow.Text = "Project Builder: exibir subprojeto " & FormularioDados.txtPEP.Value Then
    Session.findById("wnd[0]/tbar[1]/btn[13]").press
    If Session.findById("wnd[0]/sbar").Text = "Não foram bloqueados todos os objetos (ver protocolo de bloqueio)" Then
        Session.findById("wnd[0]/tbar[0]/btn[15]").press
        MsgBox "PEP está aberto por outro usuário. Verificar e tentar novamente mais tarde.", vbOKOnly
        Exit Sub
    End If
Else
    If Session.findById("wnd[0]/sbar").Text = "Não foram bloqueados todos os objetos (ver protocolo de bloqueio)" Then
        Session.findById("wnd[0]/tbar[0]/btn[15]").press
        MsgBox "PEP está aberto por outro usuário. Verificar e tentar novamente mais tarde.", vbOKOnly
        Exit Sub
    End If
End If

'Verificar parâmetros da definição do projeto
Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000001"
Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCJWB:3998/tabsPTABSCR/tabpPCNT").Select
If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCJWB:3998/tabsPTABSCR/tabpPCNT/ssubSUBSCR2:SAPLCJWB:1404/cmbPROJ-SCPRF").Key <> "Z00000000002" Then
    Session.findById("wnd[0]/tbar[0]/btn[15]").press
    MsgBox "Perfil de programação da Definição do projeto incorreto. Favor verificar.", vbOKOnly
    Exit Sub
End If
If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCJWB:3998/tabsPTABSCR/tabpPCNT/ssubSUBSCR2:SAPLCJWB:1404/cmbPROJ-SCHTYP").Key <> "" Then
    Session.findById("wnd[0]/tbar[0]/btn[15]").press
    MsgBox "Cenário de programação de prazo da Definição do projeto incorreto. Favor verificar.", vbOKOnly
    Exit Sub
End If
If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCJWB:3998/tabsPTABSCR/tabpPCNT/ssubSUBSCR2:SAPLCJWB:1404/cmbPROJ-VGPLF").Key <> "3" Then
    Session.findById("wnd[0]/tbar[0]/btn[15]").press
    MsgBox "Método de planejamento de datas base da Definição do projeto incorreto. Favor verificar.", vbOKOnly
    Exit Sub
End If
If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCJWB:3998/tabsPTABSCR/tabpPCNT/ssubSUBSCR2:SAPLCJWB:1404/cmbPROJ-EWPLF").Key <> "3" Then
    Session.findById("wnd[0]/tbar[0]/btn[15]").press
    MsgBox "Método de planejamento de datas previstas da Definição do projeto incorreto. Favor verificar.", vbOKOnly
    Exit Sub
End If
Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"

'Abrir aba de tarefas e inserir tempos
Dim nr_linhas_visiveis As Integer
Dim tempos_ok As Boolean

tempos_ok = False

Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "ACTY_OVW"
Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").Columns.elementAt(2).Width = 6
nr_linhas_visiveis = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").visiblerowcount

If nr_linhas > nr_linhas_visiveis Then
    y = nr_linhas_visiveis - 1
Else
    y = nr_linhas - 1
    tempos_ok = True
End If

For x = 0 To y
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-DAUNO[2," & x & "]").Text = Range("C" & x + 4).Value
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-ARBEI[4," & x & "]").Text = Range("E" & x + 4).Value
Next
Session.findById("wnd[0]").sendVKey 0

'Passar para página 2 e inserir tempos
If tempos_ok = False Then
    If nr_linhas > (2 * nr_linhas_visiveis) Then
        y = nr_linhas_visiveis - 1
    Else
        y = nr_linhas - nr_linhas_visiveis - 1
        tempos_ok = True
    End If
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = nr_linhas_visiveis
    For x = 0 To y
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-DAUNO[2," & x & "]").Text = Range("C" & x + nr_linhas_visiveis + 4).Value
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-ARBEI[4," & x & "]").Text = Range("E" & x + nr_linhas_visiveis + 4).Value
    Next
    Session.findById("wnd[0]").sendVKey 0
End If

'Passar para página 3 e inserir tempos
If tempos_ok = False Then
    If nr_linhas > (3 * nr_linhas_visiveis) Then
        y = nr_linhas_visiveis - 1
    Else
        y = nr_linhas - (2 * nr_linhas_visiveis) - 1
        tempos_ok = True
    End If
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = 2 * nr_linhas_visiveis
    For x = 0 To y
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-DAUNO[2," & x & "]").Text = Range("C" & x + (2 * nr_linhas_visiveis) + 4).Value
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-ARBEI[4," & x & "]").Text = Range("E" & x + (2 * nr_linhas_visiveis) + 4).Value
    Next
    Session.findById("wnd[0]").sendVKey 0
End If

'Passar para página 4 e inserir tempos
If tempos_ok = False Then
    If nr_linhas > (4 * nr_linhas_visiveis) Then
        y = nr_linhas_visiveis - 1
    Else
        y = nr_linhas - (3 * nr_linhas_visiveis) - 1
        tempos_ok = True
    End If
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = 3 * nr_linhas_visiveis
    For x = 0 To y
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-DAUNO[2," & x & "]").Text = Range("C" & x + (3 * nr_linhas_visiveis) + 4).Value
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-ARBEI[4," & x & "]").Text = Range("E" & x + (3 * nr_linhas_visiveis) + 4).Value
    Next
    Session.findById("wnd[0]").sendVKey 0
End If

'Passar para página 5 e inserir tempos
If tempos_ok = False Then
    If nr_linhas > (5 * nr_linhas_visiveis) Then
        y = nr_linhas_visiveis - 1
    Else
        y = nr_linhas - (4 * nr_linhas_visiveis) - 1
        tempos_ok = True
    End If
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = 4 * nr_linhas_visiveis
    For x = 0 To y
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-DAUNO[2," & x & "]").Text = Range("C" & x + (4 * nr_linhas_visiveis) + 4).Value
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-ARBEI[4," & x & "]").Text = Range("E" & x + (4 * nr_linhas_visiveis) + 4).Value
    Next
    Session.findById("wnd[0]").sendVKey 0
End If

'Passar para página 6 e inserir tempos
If tempos_ok = False Then
    If nr_linhas > (6 * nr_linhas_visiveis) Then
        y = nr_linhas_visiveis - 1
    Else
        y = nr_linhas - (5 * nr_linhas_visiveis) - 1
        tempos_ok = True
    End If
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = 5 * nr_linhas_visiveis
    For x = 0 To y
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-DAUNO[2," & x & "]").Text = Range("C" & x + (5 * nr_linhas_visiveis) + 4).Value
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-ARBEI[4," & x & "]").Text = Range("E" & x + (5 * nr_linhas_visiveis) + 4).Value
    Next
    Session.findById("wnd[0]").sendVKey 0
End If

'Passar para página 7 e inserir tempos
If tempos_ok = False Then
    If nr_linhas > (7 * nr_linhas_visiveis) Then
        y = nr_linhas_visiveis - 1
    Else
        y = nr_linhas - (6 * nr_linhas_visiveis) - 1
        tempos_ok = True
    End If
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = 6 * nr_linhas_visiveis
    For x = 0 To y
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-DAUNO[2," & x & "]").Text = Range("C" & x + (6 * nr_linhas_visiveis) + 4).Value
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-ARBEI[4," & x & "]").Text = Range("E" & x + (6 * nr_linhas_visiveis) + 4).Value
    Next
    Session.findById("wnd[0]").sendVKey 0
End If

'Passar para página 8 e inserir tempos
If tempos_ok = False Then
    If nr_linhas > (8 * nr_linhas_visiveis) Then
        y = nr_linhas_visiveis - 1
    Else
        y = nr_linhas - (7 * nr_linhas_visiveis) - 1
        tempos_ok = True
    End If
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = 7 * nr_linhas_visiveis
    For x = 0 To y
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-DAUNO[2," & x & "]").Text = Range("C" & x + (7 * nr_linhas_visiveis) + 4).Value
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-ARBEI[4," & x & "]").Text = Range("E" & x + (7 * nr_linhas_visiveis) + 4).Value
    Next
    Session.findById("wnd[0]").sendVKey 0
End If

'Passar para página 9 e inserir tempos
If tempos_ok = False Then
    If nr_linhas > (9 * nr_linhas_visiveis) Then
        y = nr_linhas_visiveis - 1
    Else
        y = nr_linhas - (8 * nr_linhas_visiveis) - 1
        tempos_ok = True
    End If
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = 8 * nr_linhas_visiveis
    For x = 0 To y
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-DAUNO[2," & x & "]").Text = Range("C" & x + (8 * nr_linhas_visiveis) + 4).Value
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-ARBEI[4," & x & "]").Text = Range("E" & x + (8 * nr_linhas_visiveis) + 4).Value
    Next
    Session.findById("wnd[0]").sendVKey 0
End If

'Passar para página 9 e inserir tempos
If tempos_ok = False Then
    If nr_linhas > (10 * nr_linhas_visiveis) Then
        y = nr_linhas_visiveis - 1
    Else
        y = nr_linhas - (9 * nr_linhas_visiveis) - 1
        tempos_ok = True
    End If
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = 9 * nr_linhas_visiveis
    For x = 0 To y
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-DAUNO[2," & x & "]").Text = Range("C" & x + (9 * nr_linhas_visiveis) + 4).Value
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-ARBEI[4," & x & "]").Text = Range("E" & x + (9 * nr_linhas_visiveis) + 4).Value
    Next
    Session.findById("wnd[0]").sendVKey 0
End If

    Session.findById("wnd[0]").sendVKey 0
    Session.findById("wnd[0]").sendVKey 0

      'Salvar CJ20N
                                
    Session.findById("wnd[0]/tbar[0]/btn[11]").press
                                
       'Aviso orçamento de custos CPC17
     If Session.ActiveWindow.Text = "Ctrl.disponibilidade" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
                                
    'Erro na determinação dos custos
    If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
    If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
    If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
    If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
    If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
    If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
    If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
    If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
    If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
    If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
    If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
    If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
    If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
    If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
    If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
    If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
    If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
    If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
    If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
                            
    'Aviso Bloco K
    If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
    If Session.ActiveWindow.Text = "Baixa existente em período anterior" Then Session.findById("wnd[1]/usr/btnBUTTON_1").press

    If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
    If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
    If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
    If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
    If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
    If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
    If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press

'Sair da CJ20N
Session.findById("wnd[0]/tbar[0]/btn[15]").press
                                
                        
'Abrir CJ20N e inserir PEP
Session.findById("wnd[0]").maximize
Session.findById("wnd[0]/tbar[0]/okcd").Text = "CJ20N"
Session.findById("wnd[0]").sendVKey 0
Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[0]/shell").pressButton "OPEN"
Session.findById("wnd[1]/usr/ctxtCNPB_W_ADD_OBJ_DYN-PROJ_EXT").Text = ""
Session.findById("wnd[1]/usr/ctxtCNPB_W_ADD_OBJ_DYN-PRPS_EXT").Text = FormularioDados.txtPEP.Value
Session.findById("wnd[1]/tbar[0]/btn[0]").press
                                               
'Inserir data início do projeto e planejar

If Session.findById("wnd[0]/sbar").Text = "Não foram bloqueados todos os objetos (ver protocolo de bloqueio)" Then Session.findById("wnd[0]/tbar[1]/btn[13]").press
Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_DETAIL/shellcont/shell").pressButton "WBSE_DET"
Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCJWB:3999/tabsTABCJWB/tabpTERM").Select
Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCJWB:3999/tabsTABCJWB/tabpTERM/ssubSUBSCR1:SAPLCJTR:3500/ctxtPRTE-PSTRT").Text = FormularioDados.txtDataInicio.Value
Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCJWB:3999/tabsTABCJWB/tabpTERM/ssubSUBSCR1:SAPLCJTR:3500/txtPRTE-PDAUR").Text = ""
Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCJWB:3999/tabsTABCJWB/tabpTERM/ssubSUBSCR1:SAPLCJTR:3500/ctxtPRTE-PENDE").Text = ""
Session.findById("wnd[0]").sendVKey 0
Session.findById("wnd[0]").sendVKey 33

'Liberar tarefa inicial
If FormularioDados.chkCPC17.Value = False Then
    If FormularioDados.tipoestrutura.Value <> "ESSW (elétrica)" Then
        Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "ACTY_OVW"
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = 0
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(0).Selected = True
        If FormularioDados.SemEng.Value = True And FormularioDados.tipoestrutura.Value = "Container Solar" And FormularioDados.Betim1310.Value = False Then
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(1).Selected = True
        End If
        Session.findById("wnd[0]").sendVKey 26
    End If
End If


'Salvar
Session.findById("wnd[0]/tbar[0]/btn[11]").press

'Aviso orçamento de custos CPC17
If Session.ActiveWindow.Text = "Ctrl.disponibilidade" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press

'Erro na determinação dos custos
If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press


'Aviso Bloco K
If Session.ActiveWindow.Text = "Baixa existente em período anterior" Then Session.findById("wnd[1]/usr/btnBUTTON_1").press

'Sair da CJ20N
Session.findById("wnd[0]/tbar[0]/btn[15]").press
                                                                                              

End Sub
```

#### `campos_usuario` (Function)

```vb
Function campos_usuario()
    
' Espera porque a rede é lenta
Application.Wait (Now + TimeValue("0:00:05"))

Call campos_usuario1

End Function
```

#### `campos_usuario1` (Sub)

```vb
Sub campos_usuario1()

Dim SapGuiAuto As Object
Dim Application As Object
Dim Connection As Object
Dim Session As Object
Dim WScript As Object
Dim j As Long
Dim sessao_ok As Boolean
Dim Nomenc As String
Dim numero_modulos As String
Dim x As Integer
Dim y As Integer

'Verificar número de módulos
If FormularioDados.nrmodulos.Value = "1 Módulo" Then numero_modulos = "1"
If FormularioDados.nrmodulos.Value = "2 Módulos" Then numero_modulos = "2"
If FormularioDados.nrmodulos.Value = "3 Módulos" Then numero_modulos = "3"
If FormularioDados.nrmodulos.Value = "4 Módulos" Then numero_modulos = "4"
If FormularioDados.nrmodulos.Value = "5 Módulos" Then numero_modulos = "5"
If FormularioDados.nrmodulos.Value = "6 Módulos" Then numero_modulos = "6"
If FormularioDados.nrmodulos.Value = "7 Módulos" Then numero_modulos = "7"
If FormularioDados.nrmodulos.Value = "8 Módulos" Then numero_modulos = "8"

'Conexão com o Objeto SAP
Set SapGuiAuto = GetObject("SAPGUI")
Set Application = SapGuiAuto.GetScriptingEngine
Set Connection = Application.Children(0)

sessao_ok = False
For j = 0 To Application.Children(0).Sessions.Count() - 1
    Set Session = Connection.Children(CLng(j))
    If Session.ActiveWindow.Text = "SAP Easy Access" Then
        sessao_ok = True
        Exit For
    End If
Next

If sessao_ok = False Then
    MsgBox "Nenhuma janela do SAP na tela inicial foi encontrada. Programa interrompido.", vbOKOnly
    Exit Sub
End If

ThisWorkbook.Application.StatusBar = "PEP " & FormularioDados.txtPEP.Value & ": Preenchendo campos usuário"

Nomenc = Range("F3").Value

'Abrir CJ20N e inserir PEP
Session.findById("wnd[0]").maximize
Session.findById("wnd[0]/tbar[0]/okcd").Text = "/NCJ20N"
Session.findById("wnd[0]").sendVKey 0
Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[0]/shell").pressButton "OPEN"
Session.findById("wnd[1]/usr/ctxtCNPB_W_ADD_OBJ_DYN-PROJ_EXT").Text = ""
Session.findById("wnd[1]/usr/ctxtCNPB_W_ADD_OBJ_DYN-PRPS_EXT").Text = FormularioDados.txtPEP.Value
Session.findById("wnd[1]").sendVKey 0

'Encerrar programa em caso de mensagem de erro (PEP inexistente)
If Session.ActiveWindow.Text = "Erro" Then
    Session.findById("wnd[2]").sendVKey 0
    Session.findById("wnd[1]").Close
    Session.findById("wnd[0]/tbar[0]/btn[15]").press
    MsgBox "PEP informado não existe. Programa interrompido.", vbOKOnly
    Exit Sub
End If

'Por em modo de edição
If Session.ActiveWindow.Text = "Project Builder: exibir subprojeto " & FormularioDados.txtPEP.Value Then
    Session.findById("wnd[0]/tbar[1]/btn[13]").press
    If Session.findById("wnd[0]/sbar").Text = "Não foram bloqueados todos os objetos (ver protocolo de bloqueio)" Then
        Session.findById("wnd[0]/tbar[0]/btn[15]").press
        MsgBox "PEP está aberto por outro usuário. Verificar e tentar novamente mais tarde.", vbOKOnly
        Exit Sub
    End If
Else
    If Session.findById("wnd[0]/sbar").Text = "Não foram bloqueados todos os objetos (ver protocolo de bloqueio)" Then
        Session.findById("wnd[0]/tbar[0]/btn[15]").press
        MsgBox "PEP está aberto por outro usuário. Verificar e tentar novamente mais tarde.", vbOKOnly
        Exit Sub
    End If
End If

'Verificar número de diagramas
If FormularioDados.tipoestrutura.Value = "ESSW (elétrica)" Then y = 0
If FormularioDados.tipoestrutura.Value = "Pilotis" Then y = 1
If FormularioDados.tipoestrutura.Value = "Skid (mecânica)" Then y = 2
If FormularioDados.tipoestrutura.Value = "Skid (mecânica)" And FormularioDados.Betim1310.Value = True Then y = 3
If FormularioDados.tipoestrutura.Value = "Skid (com elétrica)" And FormularioDados.Betim1310.Value = True Then y = 4
If FormularioDados.nrmodulos.Value = "1 Módulo" And FormularioDados.tipoestrutura.Value <> "Skid (mecânica)" And FormularioDados.tipoestrutura.Value <> "Pilotis" And FormularioDados.tipoestrutura.Value <> "ESSW (elétrica)" Then y = 4
If FormularioDados.nrmodulos.Value = "2 Módulos" Then y = 5
If FormularioDados.nrmodulos.Value = "3 Módulos" Then y = 6
If FormularioDados.nrmodulos.Value = "4 Módulos" Then y = 7
If FormularioDados.nrmodulos.Value = "5 Módulos" Then y = 8
If FormularioDados.nrmodulos.Value = "6 Módulos" Then y = 9
If FormularioDados.nrmodulos.Value = "7 Módulos" Then y = 10
If FormularioDados.nrmodulos.Value = "8 Módulos" Then y = 11
If FormularioDados.proBTI.Value = True Then y = 4
If FormularioDados.proBTI.Value = True And FormularioDados.SemEng.Value = False Then y = 3
If FormularioDados.SemEng.Value = True And FormularioDados.proBTI.Value = False Then y = 2
'Somar diagrama de encerramento
y = y + 1
'Executar modificação em massa
For x = 0 To y
    Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "NETW_OVW"
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtNETW_OVW-AUFNR[0," & x & "]").SetFocus
    Session.findById("wnd[0]").sendVKey 2
    Session.findById("wnd[0]/tbar[1]/btn[21]").press
    Session.findById("wnd[1]/usr/subTABSTRIP:SAPLMASSINTERFACE:0118/subTABSTRIP:SAPLMASSINTERFACE:0120/tabsTBSTRP_TABLES/tabpTAB5").Select
    
    'Selecionar campos
    Session.findById("wnd[2]/usr/ssubRAHMEN:SAPLCNFA:0111/subALLE_FELDER:SAPLCNFA:0130/btnSUCHEN").press
    Session.findById("wnd[3]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").Text = "campo de usuário"
    Session.findById("wnd[3]").sendVKey 0
    Session.findById("wnd[2]/usr/ssubRAHMEN:SAPLCNFA:0111/subALLE_FELDER:SAPLCNFA:0130/tblSAPLCNFATC_ALLE_FELDER").getAbsoluteRow(3).Selected = True
    Session.findById("wnd[2]/usr/ssubRAHMEN:SAPLCNFA:0111/subAUSWAHL:SAPLCNFA:0140/btnAUSWAEHLEN").press
    Session.findById("wnd[2]/usr/ssubRAHMEN:SAPLCNFA:0111/subALLE_FELDER:SAPLCNFA:0130/tblSAPLCNFATC_ALLE_FELDER").getAbsoluteRow(3).Selected = True
    Session.findById("wnd[2]/usr/ssubRAHMEN:SAPLCNFA:0111/subAUSWAHL:SAPLCNFA:0140/btnAUSWAEHLEN").press
    Session.findById("wnd[2]/usr/ssubRAHMEN:SAPLCNFA:0111/subALLE_FELDER:SAPLCNFA:0130/tblSAPLCNFATC_ALLE_FELDER").getAbsoluteRow(3).Selected = True
    Session.findById("wnd[2]/usr/ssubRAHMEN:SAPLCNFA:0111/subAUSWAHL:SAPLCNFA:0140/btnAUSWAEHLEN").press
    Session.findById("wnd[2]/usr/ssubRAHMEN:SAPLCNFA:0111/subALLE_FELDER:SAPLCNFA:0130/tblSAPLCNFATC_ALLE_FELDER").getAbsoluteRow(3).Selected = True
    Session.findById("wnd[2]/usr/ssubRAHMEN:SAPLCNFA:0111/subAUSWAHL:SAPLCNFA:0140/btnAUSWAEHLEN").press
    Session.findById("wnd[2]/usr/ssubRAHMEN:SAPLCNFA:0111/subALLE_FELDER:SAPLCNFA:0130/tblSAPLCNFATC_ALLE_FELDER").getAbsoluteRow(3).Selected = True
    Session.findById("wnd[2]/usr/ssubRAHMEN:SAPLCNFA:0111/subAUSWAHL:SAPLCNFA:0140/btnAUSWAEHLEN").press
    Session.findById("wnd[2]/usr/ssubRAHMEN:SAPLCNFA:0111/subALLE_FELDER:SAPLCNFA:0130/btnSUCHEN").press
    Session.findById("wnd[3]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").Text = "CAMPO usuário qtd."
    Session.findById("wnd[3]").sendVKey 0
    Session.findById("wnd[2]/usr/ssubRAHMEN:SAPLCNFA:0111/subALLE_FELDER:SAPLCNFA:0130/tblSAPLCNFATC_ALLE_FELDER").getAbsoluteRow(4).Selected = True
    Session.findById("wnd[2]/usr/ssubRAHMEN:SAPLCNFA:0111/subAUSWAHL:SAPLCNFA:0140/btnAUSWAEHLEN").press
    Session.findById("wnd[2]/usr/ssubRAHMEN:SAPLCNFA:0111/subALLE_FELDER:SAPLCNFA:0130/tblSAPLCNFATC_ALLE_FELDER").getAbsoluteRow(4).Selected = True
    Session.findById("wnd[2]/usr/ssubRAHMEN:SAPLCNFA:0111/subAUSWAHL:SAPLCNFA:0140/btnAUSWAEHLEN").press
    Session.findById("wnd[2]/usr/ssubRAHMEN:SAPLCNFA:0111/subALLE_FELDER:SAPLCNFA:0130/btnSUCHEN").press
    Session.findById("wnd[3]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").Text = "CAMPO USUÁRIO valor"
    Session.findById("wnd[3]").sendVKey 0
    Session.findById("wnd[2]/usr/ssubRAHMEN:SAPLCNFA:0111/subALLE_FELDER:SAPLCNFA:0130/tblSAPLCNFATC_ALLE_FELDER").getAbsoluteRow(9).Selected = True
    Session.findById("wnd[2]/usr/ssubRAHMEN:SAPLCNFA:0111/subAUSWAHL:SAPLCNFA:0140/btnAUSWAEHLEN").press
    Session.findById("wnd[2]/tbar[0]/btn[0]").press
    
    'Inserir informações
    Session.findById("wnd[1]/usr/subTABSTRIP:SAPLMASSINTERFACE:0118/subTABSTRIP:SAPLMASSINTERFACE:0120/tabsTBSTRP_TABLES/tabpTAB5/ssubFIELDS:SAPLMASSINTERFACE:0130/sub:SAPLMASSINTERFACE:0130/txtMOD_FIELD-VALUE-LEFT[0,40]").Text = FormularioDados.cmbPlanejador.Value
    'If FormularioDados.planpin.Value <> "Não aplicável" Then
    '    session.findById("wnd[1]/usr/subTABSTRIP:SAPLMASSINTERFACE:0118/subTABSTRIP:SAPLMASSINTERFACE:0120/tabsTBSTRP_TABLES/tabpTAB5/ssubFIELDS:SAPLMASSINTERFACE:0130/sub:SAPLMASSINTERFACE:0130/txtMOD_FIELD-VALUE-LEFT[2,40]").Text = FormularioDados.planpin.Value
    'End If
    Session.findById("wnd[1]/usr/subTABSTRIP:SAPLMASSINTERFACE:0118/subTABSTRIP:SAPLMASSINTERFACE:0120/tabsTBSTRP_TABLES/tabpTAB5/ssubFIELDS:SAPLMASSINTERFACE:0130/sub:SAPLMASSINTERFACE:0130/txtMOD_FIELD-VALUE-LEFT[4,40]").Text = FormularioDados.txtCliente.Value
    Session.findById("wnd[1]/usr/subTABSTRIP:SAPLMASSINTERFACE:0118/subTABSTRIP:SAPLMASSINTERFACE:0120/tabsTBSTRP_TABLES/tabpTAB5/ssubFIELDS:SAPLMASSINTERFACE:0130/sub:SAPLMASSINTERFACE:0130/txtMOD_FIELD-VALUE-LEFT[6,40]").Text = Nomenc
    Session.findById("wnd[1]/usr/subTABSTRIP:SAPLMASSINTERFACE:0118/subTABSTRIP:SAPLMASSINTERFACE:0120/tabsTBSTRP_TABLES/tabpTAB5/ssubFIELDS:SAPLMASSINTERFACE:0130/sub:SAPLMASSINTERFACE:0130/ctxtMOD_FIELD-VALUE-LEFT[8,40]").Text = FormularioDados.txtDataOV.Value
    Session.findById("wnd[1]/usr/subTABSTRIP:SAPLMASSINTERFACE:0118/subTABSTRIP:SAPLMASSINTERFACE:0120/tabsTBSTRP_TABLES/tabpTAB5/ssubFIELDS:SAPLMASSINTERFACE:0130/sub:SAPLMASSINTERFACE:0130/txtMOD_FIELD-VALUE-LEFT[10,40]").Text = numero_modulos
    Session.findById("wnd[1]/usr/subTABSTRIP:SAPLMASSINTERFACE:0118/subTABSTRIP:SAPLMASSINTERFACE:0120/tabsTBSTRP_TABLES/tabpTAB5/ssubFIELDS:SAPLMASSINTERFACE:0130/sub:SAPLMASSINTERFACE:0130/txtMOD_FIELD-VALUE-LEFT[12,40]").Text = "1"
    If x < y - 1 Then
        Session.findById("wnd[1]/usr/subTABSTRIP:SAPLMASSINTERFACE:0118/subTABSTRIP:SAPLMASSINTERFACE:0120/tabsTBSTRP_TABLES/tabpTAB5/ssubFIELDS:SAPLMASSINTERFACE:0130/sub:SAPLMASSINTERFACE:0130/txtMOD_FIELD-VALUE-LEFT[14,40]").Text = FormularioDados.txtValor.Value
    Else
        If FormularioDados.tipoestrutura.Value <> "ESSW (mecânica)" And FormularioDados.tipoestrutura.Value <> "Skid (mecânica)" And FormularioDados.tipoestrutura.Value <> "Pilotis" Then
            Session.findById("wnd[1]/usr/subTABSTRIP:SAPLMASSINTERFACE:0118/subTABSTRIP:SAPLMASSINTERFACE:0120/tabsTBSTRP_TABLES/tabpTAB5/ssubFIELDS:SAPLMASSINTERFACE:0130/sub:SAPLMASSINTERFACE:0130/txtMOD_FIELD-VALUE-LEFT[14,40]").Text = FormularioDados.txtValorEletr.Value
        Else
            Session.findById("wnd[1]/usr/subTABSTRIP:SAPLMASSINTERFACE:0118/subTABSTRIP:SAPLMASSINTERFACE:0120/tabsTBSTRP_TABLES/tabpTAB5/ssubFIELDS:SAPLMASSINTERFACE:0130/sub:SAPLMASSINTERFACE:0130/txtMOD_FIELD-VALUE-LEFT[14,40]").Text = FormularioDados.txtValor.Value
        End If
    End If
    

    'Executar e fechar tela de modificação em massa
    Session.findById("wnd[1]/tbar[0]/btn[8]").press
    If Session.ActiveWindow.Text = "Informação" Then Session.findById("wnd[1]/tbar[0]/btn[0]").press
    If Session.ActiveWindow.Text = "Atualização em massa: visão de campos" Then Session.findById("wnd[1]").Close
    Session.findById("wnd[0]/tbar[0]/btn[3]").press
    If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]").Close
    
    Next
    
    'Selecionar PEP
    Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCJWB:3999/tabsTABCJWB/tabpUSR1").Select
    
    'Inserir informações no PEP
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCJWB:3999/tabsTABCJWB/tabpUSR1/ssubSUBSCR1:SAPLCJWB:1460/ctxtPRPS-USR00").Text = FormularioDados.txtCliente.Value
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCJWB:3999/tabsTABCJWB/tabpUSR1/ssubSUBSCR1:SAPLCJWB:1460/ctxtPRPS-USR01").Text = Nomenc
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCJWB:3999/tabsTABCJWB/tabpUSR1/ssubSUBSCR1:SAPLCJWB:1460/ctxtPRPS-USR02").Text = FormularioDados.cmbPlanejador.Value
    'If FormularioDados.planpin.Value <> "Não aplicável" Then
    '    session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCJWB:3999/tabsTABCJWB/tabpUSR1/ssubSUBSCR1:SAPLCJWB:1460/ctxtPRPS-USR03").Text = FormularioDados.planpin.Value
    'End If
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCJWB:3999/tabsTABCJWB/tabpUSR1/ssubSUBSCR1:SAPLCJWB:1460/ctxtPRPS-USR04").Text = numero_modulos
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCJWB:3999/tabsTABCJWB/tabpUSR1/ssubSUBSCR1:SAPLCJWB:1460/ctxtPRPS-USR05").Text = "1"
    If FormularioDados.tipoestrutura.Value <> "ESSW (mecânica)" And FormularioDados.tipoestrutura.Value <> "Skid (mecânica)" And FormularioDados.tipoestrutura.Value <> "Pilotis" Then
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCJWB:3999/tabsTABCJWB/tabpUSR1/ssubSUBSCR1:SAPLCJWB:1460/ctxtPRPS-USR07").Text = FormularioDados.txtValorEletr.Value
    Else
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCJWB:3999/tabsTABCJWB/tabpUSR1/ssubSUBSCR1:SAPLCJWB:1460/ctxtPRPS-USR07").Text = FormularioDados.txtValor.Value
    End If
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCJWB:3999/tabsTABCJWB/tabpUSR1/ssubSUBSCR1:SAPLCJWB:1460/ctxtPRPS-USE07").Text = "BRL"
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCJWB:3999/tabsTABCJWB/tabpUSR1/ssubSUBSCR1:SAPLCJWB:1460/ctxtPRPS-USR08").Text = FormularioDados.txtDataOV.Value
    
    'PREENCHE NÚMERO E ITEM DA OV
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCJWB:3999/tabsTABCJWB/tabp+CUE").Select
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCJWB:3999/tabsTABCJWB/tabp+CUE/ssubSUBSCR1:SAPLCJWB:1215/ssubCUSTSCR1:SAPLXCN1:0700/ctxtPRPS-ZZVBELN").Text = FormularioDados.txt_N_OV.Value
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCJWB:3999/tabsTABCJWB/tabp+CUE/ssubSUBSCR1:SAPLCJWB:1215/ssubCUSTSCR1:SAPLXCN1:0700/txtPRPS-ZZPOSNR").Text = FormularioDados.txt_itemOV.Value
    
    Session.findById("wnd[0]").sendVKey 0
    
    'Chama macro de horas orçadas
    Call Horas_Orçadas
    
    'Salvar
    Session.findById("wnd[0]/tbar[0]/btn[11]").press
    
    ' Aviso agr de diagramas
    If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
    If Session.ActiveWindow.Text = "Agrupamento de Diagramas" Then Session.findById("wnd[1]/usr/btnBUTTON_2").press
    If Session.ActiveWindow.Text = "Agrupamento de Diagramas" Then Session.findById("wnd[1]/usr/btnBUTTON_2").press
    If Session.ActiveWindow.Text = "Agrupamento de Diagramas" Then Session.findById("wnd[1]/usr/btnBUTTON_2").press
    If Session.ActiveWindow.Text = "Agrupamento de Diagramas" Then Session.findById("wnd[1]/usr/btnBUTTON_2").press
    If Session.ActiveWindow.Text = "Agrupamento de Diagramas" Then Session.findById("wnd[1]/usr/btnBUTTON_2").press
    If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
    If Session.ActiveWindow.Text = "Agrupamento de Diagramas" Then Session.findById("wnd[1]/usr/btnBUTTON_2").press
    If Session.ActiveWindow.Text = "Agrupamento de Diagramas" Then Session.findById("wnd[1]/usr/btnBUTTON_2").press
    If Session.ActiveWindow.Text = "Agrupamento de Diagramas" Then Session.findById("wnd[1]/usr/btnBUTTON_2").press
    If Session.ActiveWindow.Text = "Agrupamento de Diagramas" Then Session.findById("wnd[1]/usr/btnBUTTON_2").press
    If Session.ActiveWindow.Text = "Agrupamento de Diagramas" Then Session.findById("wnd[1]/usr/btnBUTTON_2").press
    If Session.ActiveWindow.Text = "Agrupamento de Diagramas" Then Session.findById("wnd[1]/usr/btnBUTTON_2").press
    If Session.ActiveWindow.Text = "Agrupamento de Diagramas" Then Session.findById("wnd[1]/usr/btnBUTTON_2").press
    If Session.ActiveWindow.Text = "Agrupamento de Diagramas" Then Session.findById("wnd[1]/usr/btnBUTTON_2").press
    
    'Aviso orçamento de custos CPC47
    If Session.ActiveWindow.Text = "Ctrl.disponibilidade" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
    
    'Erro na determinação dos custos
    If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
    If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
    If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
    If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
    If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
    If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
    If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
    If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
    If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
    If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
    If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
    If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
    If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press

   
    'Aviso Bloco K
    If Session.ActiveWindow.Text = "Baixa existente em período anterior" Then Session.findById("wnd[1]/usr/btnBUTTON_1").press
    
    'Sair da CJ20N
    Session.findById("wnd[0]/tbar[0]/btn[15]").press


'Abrir CJ20N e inserir PEP SELECIONAR DIAGRAMAS

    If FormularioDados.tipoestrutura.Value = "Container Solar" And FormularioDados.SemEng.Value = True Then
    
        Session.findById("wnd[0]").maximize
        Session.findById("wnd[0]/tbar[0]/okcd").Text = "/NCJ20N"
        Session.findById("wnd[0]").sendVKey 0
        Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[0]/shell").pressButton "OPEN"
        Session.findById("wnd[1]/usr/ctxtCNPB_W_ADD_OBJ_DYN-PROJ_EXT").Text = ""
        Session.findById("wnd[1]/usr/ctxtCNPB_W_ADD_OBJ_DYN-PRPS_EXT").Text = FormularioDados.txtPEP.Value
        Session.findById("wnd[1]").sendVKey 0
        
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "NETW_OVW"
        MEC1WIJ = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtNETW_OVW-AUFNR[0,0]").Text
        ACESWIJ = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtNETW_OVW-AUFNR[0,1]").Text
        ELE1WIJ = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtNETW_OVW-AUFNR[0,2]").Text
        PROBTI1 = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtNETW_OVW-AUFNR[0,3]").Text
        
        'sair
        Session.findById("wnd[0]/tbar[0]/btn[3]").press
        'Sair da CJ20N
        Session.findById("wnd[0]/tbar[0]/btn[15]").press
    
    End If


                        
' Liberar tarefas caso não tenha ENG SKID

'    If FormularioDados.SemEng.Value = True And FormularioDados.tipoestrutura.Value = "Skid (com elétrica)" And FormularioDados.Betim1310.Value = True And FormularioDados.chkCPC17.Value = False Then
'
'          ' Abrir CN47N
'           session.findById("wnd[0]/tbar[0]/okcd").Text = "/nCN47N"
'           session.findById("wnd[0]").sendVKey 0
'            'Preencher perfil caso necessário
'                If session.ActiveWindow.Text = "Entrar perfil" Then
'                    session.findById("wnd[1]/usr/ctxtTCNT-PROF_DB").Text = "000000000001"
'                    session.findById("wnd[1]/tbar[0]/btn[0]").press
'                    End If
'                            session.findById("wnd[0]/usr/ctxtCN_PSPNR-LOW").Text = FormularioDados.txtPEP.Value
'                            session.findById("wnd[0]/usr/ctxtP_DISVAR").Text = ""
'                            session.findById("wnd[0]/usr/ctxtCN_NETNR-LOW").Text = ""
'                            session.findById("wnd[0]/usr/btn%_CN_ACTVT_%_APP_%-VALU_PUSH").press
'                            session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "710"
'                            session.findById("wnd[1]/tbar[0]/btn[8]").press
'                            session.findById("wnd[0]/usr/ctxtP_DISVAR").Text = "/WAU-PCP_K"
'                            session.findById("wnd[0]/tbar[1]/btn[8]").press
'
'                            session.findById("wnd[0]/usr/cntlALVCONTAINER/shellcont/shell").doubleClickCurrentCell
'                            session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").topNode = "000001"
'                            session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell").topNode = "         23"
'                            session.findById("wnd[0]/tbar[1]/btn[13]").press
'
'                            session.findById("wnd[0]").sendVKey 26
'                            'Salvar
'                            session.findById("wnd[0]/tbar[0]/btn[11]").press
'
'                            'Aviso orçamento de custos CPC17
'                            If session.ActiveWindow.Text = "Ctrl.disponibilidade" Then session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
'
'                            'Erro na determinação dos custos
'                            If session.ActiveWindow.Text = "Determ.custos" Then session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
'
'                            'Aviso Bloco K
'                            If session.ActiveWindow.Text = "Baixa existente em período anterior" Then session.findById("wnd[1]/usr/btnBUTTON_1").press
'
'                            'Voltar ao inicio
'                            session.findById("wnd[0]/tbar[0]/btn[3]").press
'                            session.findById("wnd[0]/tbar[0]/btn[3]").press
'
'
'    End If



    'Ligação
    If FormularioDados.tipoestrutura.Value = "Container Solar" And FormularioDados.SemEng.Value = True And FormularioDados.proBTI.Value = True Then
        ' Abrir CN47N
           Session.findById("wnd[0]/tbar[0]/okcd").Text = "/nCN47N"
           Session.findById("wnd[0]").sendVKey 0
            'Preencher perfil caso necessário
                If Session.ActiveWindow.Text = "Entrar perfil" Then
                    Session.findById("wnd[1]/usr/ctxtTCNT-PROF_DB").Text = "000000000001"
                    Session.findById("wnd[1]/tbar[0]/btn[0]").press
                    End If
                            Session.findById("wnd[0]/usr/ctxtCN_PSPNR-LOW").Text = FormularioDados.txtPEP.Value
                            Session.findById("wnd[0]/usr/ctxtP_DISVAR").Text = ""
                            Session.findById("wnd[0]/usr/ctxtCN_NETNR-LOW").Text = ""
                            Session.findById("wnd[0]/usr/btn%_CN_ACTVT_%_APP_%-VALU_PUSH").press
                            Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "780"
                            Session.findById("wnd[1]/tbar[0]/btn[8]").press
                            Session.findById("wnd[0]/usr/ctxtP_DISVAR").Text = "/WAU-PCP_K"
                            Session.findById("wnd[0]/tbar[1]/btn[8]").press
                            
                                Session.findById("wnd[0]/usr/cntlALVCONTAINER/shellcont/shell").doubleClickCurrentCell
                                Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").topNode = "000001"
                                Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell").topNode = "         23"
                                Session.findById("wnd[0]/tbar[1]/btn[13]").press
                                Session.findById("wnd[0]/mbar/menu[4]/menu[8]").Select
                                Session.findById("wnd[1]/usr/chkCNPB_OPTIONS-RELATION").Selected = True
                                Session.findById("wnd[1]/usr/chkCNPB_OPTIONS-RELATION").SetFocus
                                Session.findById("wnd[1]/tbar[0]/btn[0]").press
                                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "RELA_OVW"
                                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2403/tabsTABSTRIP_2401/tabpAOBG/ssubSUBSCR_2401:SAPLCOVG:2404/tblSAPLCOVGTCTRL_2402/ctxtAFABD-VORNR[0,5]").Text = "730"
                                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2403/tabsTABSTRIP_2401/tabpAOBG/ssubSUBSCR_2401:SAPLCOVG:2404/tblSAPLCOVGTCTRL_2402/ctxtAFABD-NETZPLAN[1,5]").Text = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2403/tabsTABSTRIP_2401/tabpAOBG/ssubSUBSCR_2401:SAPLCOVG:2404/tblSAPLCOVGTCTRL_2402/ctxtAFABD-NETZPLAN[1,1]").Text - 1
                                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2403/tabsTABSTRIP_2401/tabpAOBG/ssubSUBSCR_2401:SAPLCOVG:2404/tblSAPLCOVGTCTRL_2402/txtAFABD-DAUER[4,5]").Text = "15"
                                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2403/tabsTABSTRIP_2401/tabpAOBG/ssubSUBSCR_2401:SAPLCOVG:2404/tblSAPLCOVGTCTRL_2402/ctxtAFABD-ZEINH[5,5]").Text = "dia"
                                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2403/tabsTABSTRIP_2401/tabpAOBG/ssubSUBSCR_2401:SAPLCOVG:2404/tblSAPLCOVGTCTRL_2402/ctxtAFABD-NETZPLAN[1,5]").SetFocus
                                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2403/tabsTABSTRIP_2401/tabpAOBG/ssubSUBSCR_2401:SAPLCOVG:2404/tblSAPLCOVGTCTRL_2402/ctxtAFABD-NETZPLAN[1,5]").caretPosition = 9
                                Session.findById("wnd[0]").sendVKey 0
                                'Salvar
                                Session.findById("wnd[0]/tbar[0]/btn[11]").press
                                
                                'Aviso orçamento de custos CPC17
                                If Session.ActiveWindow.Text = "Ctrl.disponibilidade" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
                                If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
                                
                                'Erro na determinação dos custos
                                If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
                                If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
                                
                                'Aviso Bloco K
                                If Session.ActiveWindow.Text = "Baixa existente em período anterior" Then Session.findById("wnd[1]/usr/btnBUTTON_1").press
                                
                                
                                Session.findById("wnd[0]/tbar[0]/btn[3]").press
                                Session.findById("wnd[0]/tbar[0]/btn[3]").press
    End If
       
            
    If FormularioDados.tipoestrutura.Value = "Container Solar" And FormularioDados.SemEng.Value = True And FormularioDados.proBTI.Value = True Then
        ' Abrir CN47N
           Session.findById("wnd[0]/tbar[0]/okcd").Text = "/nCN47N"
           Session.findById("wnd[0]").sendVKey 0
            'Preencher perfil caso necessário
                If Session.ActiveWindow.Text = "Entrar perfil" Then
                    Session.findById("wnd[1]/usr/ctxtTCNT-PROF_DB").Text = "000000000001"
                    Session.findById("wnd[1]/tbar[0]/btn[0]").press
                    End If
                            Session.findById("wnd[0]/usr/ctxtCN_PSPNR-LOW").Text = FormularioDados.txtPEP.Value
                            Session.findById("wnd[0]/usr/ctxtP_DISVAR").Text = ""
                            Session.findById("wnd[0]/usr/ctxtCN_NETNR-LOW").Text = ""
                            Session.findById("wnd[0]/usr/btn%_CN_ACTVT_%_APP_%-VALU_PUSH").press
                            Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "780"
                            Session.findById("wnd[1]/tbar[0]/btn[8]").press
                            Session.findById("wnd[0]/usr/ctxtP_DISVAR").Text = "/WAU-PCP_K"
                            Session.findById("wnd[0]/tbar[1]/btn[8]").press
                            
                                Session.findById("wnd[0]/usr/cntlALVCONTAINER/shellcont/shell").doubleClickCurrentCell
                                Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").topNode = "000001"
                                Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell").topNode = "         23"
                                Session.findById("wnd[0]/tbar[1]/btn[13]").press
                                Session.findById("wnd[0]/mbar/menu[4]/menu[8]").Select
                                Session.findById("wnd[1]/usr/chkCNPB_OPTIONS-RELATION").Selected = True
                                Session.findById("wnd[1]/usr/chkCNPB_OPTIONS-RELATION").SetFocus
                                Session.findById("wnd[1]/tbar[0]/btn[0]").press
                                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "RELA_OVW"
                                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2403/tabsTABSTRIP_2401/tabpAOBG/ssubSUBSCR_2401:SAPLCOVG:2404/tblSAPLCOVGTCTRL_2402/ctxtAFABD-VORNR[0,4]").Text = "710"
                                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2403/tabsTABSTRIP_2401/tabpAOBG/ssubSUBSCR_2401:SAPLCOVG:2404/tblSAPLCOVGTCTRL_2402/ctxtAFABD-NETZPLAN[1,4]").Text = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2403/tabsTABSTRIP_2401/tabpAOBG/ssubSUBSCR_2401:SAPLCOVG:2404/tblSAPLCOVGTCTRL_2402/ctxtAFABD-NETZPLAN[1,0]").Text
                                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2403/tabsTABSTRIP_2401/tabpAOBG/ssubSUBSCR_2401:SAPLCOVG:2404/tblSAPLCOVGTCTRL_2402/txtAFABD-DAUER[4,4]").Text = "15"
                                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2403/tabsTABSTRIP_2401/tabpAOBG/ssubSUBSCR_2401:SAPLCOVG:2404/tblSAPLCOVGTCTRL_2402/ctxtAFABD-ZEINH[5,4]").Text = "dia"
                                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2403/tabsTABSTRIP_2401/tabpAOBG/ssubSUBSCR_2401:SAPLCOVG:2404/tblSAPLCOVGTCTRL_2402/ctxtAFABD-NETZPLAN[1,4]").SetFocus
                                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2403/tabsTABSTRIP_2401/tabpAOBG/ssubSUBSCR_2401:SAPLCOVG:2404/tblSAPLCOVGTCTRL_2402/ctxtAFABD-NETZPLAN[1,4]").caretPosition = 9
                                Session.findById("wnd[0]").sendVKey 0
                                'Salvar
                                Session.findById("wnd[0]/tbar[0]/btn[11]").press
                                
                                'Aviso orçamento de custos CPC17
                                If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
                                If Session.ActiveWindow.Text = "Ctrl.disponibilidade" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
                                
                                'Erro na determinação dos custos
                                If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
                                If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
                                
                                'Erro na determinação dos custos
                                If Session.ActiveWindow.Text = "Encerrar projeto" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
                                
                                'Aviso Bloco K
                                If Session.ActiveWindow.Text = "Baixa existente em período anterior" Then Session.findById("wnd[1]/usr/btnBUTTON_1").press
                            
                            
   Session.findById("wnd[0]/tbar[0]/btn[3]").press
   Session.findById("wnd[0]/tbar[0]/btn[3]").press
   
    End If
        
              
 ' Fazer ligações da Mecânica e de Acessórios
    If FormularioDados.tipoestrutura.Value = "Container Solar" And FormularioDados.proBTI.Value = True And FormularioDados.SemEng.Value = True Then
    ' Abrir CN47N
           Session.findById("wnd[0]/tbar[0]/okcd").Text = "/nCN47N"
           Session.findById("wnd[0]").sendVKey 0
            'Preencher perfil caso necessário
                If Session.ActiveWindow.Text = "Entrar perfil" Then
                    Session.findById("wnd[1]/usr/ctxtTCNT-PROF_DB").Text = "000000000001"
                    Session.findById("wnd[1]/tbar[0]/btn[0]").press
                    End If
                    
                            Session.findById("wnd[0]/usr/ctxtCN_PSPNR-LOW").Text = FormularioDados.txtPEP.Value
                            Session.findById("wnd[0]/usr/ctxtP_DISVAR").Text = ""
                            Session.findById("wnd[0]/usr/ctxtCN_NETNR-LOW").Text = ""
                            Session.findById("wnd[0]/usr/btn%_CN_ACTVT_%_APP_%-VALU_PUSH").press
                            Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "900"
                            Session.findById("wnd[1]/tbar[0]/btn[8]").press
                            Session.findById("wnd[0]/usr/ctxtP_DISVAR").Text = "/WAU-PCP_K"
                            Session.findById("wnd[0]/tbar[1]/btn[8]").press
                            
                            
        Session.findById("wnd[0]/usr/cntlALVCONTAINER/shellcont/shell").currentCellRow = 1
        Session.findById("wnd[0]/usr/cntlALVCONTAINER/shellcont/shell").doubleClickCurrentCell
        Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").topNode = "000001"
        Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell").topNode = "         23"
        Session.findById("wnd[0]/tbar[1]/btn[13]").press
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subIDENTIFICATION:SAPLCONW:0110/txtAFVGD-VORNR").Text = "905"
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subIDENTIFICATION:SAPLCONW:0110/txtAFVGD-VORNR").caretPosition = 3
        Session.findById("wnd[0]").sendVKey 0
        
        'Salvar
        Session.findById("wnd[0]/tbar[0]/btn[11]").press
        
        'Aviso orçamento de custos CPC17
        If Session.ActiveWindow.Text = "Ctrl.disponibilidade" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
        
        'Erro na determinação dos custos
        If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
        If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
        If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
        If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
        If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
        If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
        If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
        If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
        If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
        If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
        If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
        
        'Aviso Bloco K
        If Session.ActiveWindow.Text = "Baixa existente em período anterior" Then Session.findById("wnd[1]/usr/btnBUTTON_1").press

        Session.findById("wnd[0]/tbar[0]/btn[3]").press
        Session.findById("wnd[0]/tbar[0]/btn[3]").press
        
    End If


'Alterar valor da elétrica WIJ
If FormularioDados.tipoestrutura.Value = "Container Solar" And FormularioDados.proBTI.Value = True Then


    Session.findById("wnd[0]/tbar[0]/okcd").Text = "/NCJ20N"
    Session.findById("wnd[0]").sendVKey 0
    Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[0]/shell").pressButton "OPEN"
    Session.findById("wnd[1]/usr/ctxtCNPB_W_ADD_OBJ_DYN-PROJ_EXT").Text = ""
    Session.findById("wnd[1]/usr/ctxtCNPB_W_ADD_OBJ_DYN-PRPS_EXT").Text = FormularioDados.txtPEP.Value
    Session.findById("wnd[1]").sendVKey 0
        If Session.ActiveWindow.Text = "Project Builder: exibir subprojeto " & FormularioDados.txtPEP.Value Then
            Session.findById("wnd[0]/tbar[1]/btn[13]").press
        End If
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "NETW_OVW"
    
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtNETW_OVW-AUFNR[0,2]").SetFocus
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtNETW_OVW-AUFNR[0,2]").caretPosition = 4
    Session.findById("wnd[0]").sendVKey 2
    'session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").topNode = "000666"
    'session.findById("wnd[1]/tbar[0]/btn[0]").press
    Session.findById("wnd[0]/tbar[1]/btn[21]").press
    Session.findById("wnd[1]/usr/subTABSTRIP:SAPLMASSINTERFACE:0118/subTABSTRIP:SAPLMASSINTERFACE:0120/tabsTBSTRP_TABLES/tabpTAB5").Select
    Session.findById("wnd[2]/usr/ssubRAHMEN:SAPLCNFA:0111/subALLE_FELDER:SAPLCNFA:0130/tblSAPLCNFATC_ALLE_FELDER").verticalScrollbar.Position = 1
    Session.findById("wnd[2]/usr/ssubRAHMEN:SAPLCNFA:0111/subALLE_FELDER:SAPLCNFA:0130/tblSAPLCNFATC_ALLE_FELDER").verticalScrollbar.Position = 2
    Session.findById("wnd[2]/usr/ssubRAHMEN:SAPLCNFA:0111/subALLE_FELDER:SAPLCNFA:0130/tblSAPLCNFATC_ALLE_FELDER").verticalScrollbar.Position = 3
    Session.findById("wnd[2]/usr/ssubRAHMEN:SAPLCNFA:0111/subALLE_FELDER:SAPLCNFA:0130/tblSAPLCNFATC_ALLE_FELDER").verticalScrollbar.Position = 4
    Session.findById("wnd[2]/usr/ssubRAHMEN:SAPLCNFA:0111/subALLE_FELDER:SAPLCNFA:0130/tblSAPLCNFATC_ALLE_FELDER").verticalScrollbar.Position = 5
    Session.findById("wnd[2]/usr/ssubRAHMEN:SAPLCNFA:0111/subALLE_FELDER:SAPLCNFA:0130/tblSAPLCNFATC_ALLE_FELDER").verticalScrollbar.Position = 6
    Session.findById("wnd[2]/usr/ssubRAHMEN:SAPLCNFA:0111/subALLE_FELDER:SAPLCNFA:0130/tblSAPLCNFATC_ALLE_FELDER/txtALLE_FELDER-SCRTEXT[0,10]").SetFocus
    Session.findById("wnd[2]/usr/ssubRAHMEN:SAPLCNFA:0111/subALLE_FELDER:SAPLCNFA:0130/tblSAPLCNFATC_ALLE_FELDER/txtALLE_FELDER-SCRTEXT[0,10]").caretPosition = 18
    Session.findById("wnd[2]").sendVKey 2
    Session.findById("wnd[2]/usr/ssubRAHMEN:SAPLCNFA:0111/subALLE_FELDER:SAPLCNFA:0130/tblSAPLCNFATC_ALLE_FELDER").Columns.elementAt(0).Width = 20
    Session.findById("wnd[2]/tbar[0]/btn[0]").press
    Session.findById("wnd[1]/usr/subTABSTRIP:SAPLMASSINTERFACE:0118/subTABSTRIP:SAPLMASSINTERFACE:0120/tabsTBSTRP_TABLES/tabpTAB5/ssubFIELDS:SAPLMASSINTERFACE:0130/sub:SAPLMASSINTERFACE:0130/txtMOD_FIELD-VALUE-LEFT[0,40]").Text = FormularioDados.txtValorEletr.Value
    Session.findById("wnd[1]/usr/subTABSTRIP:SAPLMASSINTERFACE:0118/subTABSTRIP:SAPLMASSINTERFACE:0120/tabsTBSTRP_TABLES/tabpTAB5/ssubFIELDS:SAPLMASSINTERFACE:0130/sub:SAPLMASSINTERFACE:0130/txtMOD_FIELD-VALUE-LEFT[0,40]").caretPosition = 4
    Session.findById("wnd[1]/tbar[0]/btn[8]").press
    Session.findById("wnd[0]/tbar[0]/btn[3]").press

        If Session.ActiveWindow.Text = "Informação" Then
            Session.findById("wnd[1]/tbar[0]/btn[0]").press
            Session.findById("wnd[1]").Close
        End If

                        'Salvar
                        Session.findById("wnd[0]/tbar[0]/btn[11]").press
                            
                        'Aviso orçamento de custos CPC17
                        If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
                        If Session.ActiveWindow.Text = "Ctrl.disponibilidade" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
                            
                        'Erro na determinação dos custos
                        If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
                        If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
                            
                        'Aviso Bloco K
                        If Session.ActiveWindow.Text = "Baixa existente em período anterior" Then Session.findById("wnd[1]/usr/btnBUTTON_1").press
                    
                        Session.findById("wnd[0]/tbar[0]/btn[3]").press
                        'Sair da CJ20N
                        Session.findById("wnd[0]/tbar[0]/btn[15]").press

End If
   
End Sub
```

#### `amarrar_material` (Sub)

```vb
Sub amarrar_material()

Dim SapGuiAuto As Object
Dim Application As Object
Dim Connection As Object
Dim Session As Object
Dim WScript As Object
Dim j As Long
Dim sessao_ok As Boolean

'Conexão com o Objeto SAP
Set SapGuiAuto = GetObject("SAPGUI")
Set Application = SapGuiAuto.GetScriptingEngine
Set Connection = Application.Children(0)

sessao_ok = False
For j = 0 To Application.Children(0).Sessions.Count() - 1
    Set Session = Connection.Children(CLng(j))
    If Session.ActiveWindow.Text = "SAP Easy Access" Then
        sessao_ok = True
        Exit For
    End If
Next

If sessao_ok = False Then
    MsgBox "Nenhuma janela do SAP na tela inicial foi encontrada. Programa interrompido.", vbOKOnly
    Exit Sub
End If

ThisWorkbook.Application.StatusBar = "PEP " & FormularioDados.txtPEP.Value & ": Amarrando material"


'Abrir CJ20N e inserir PEP
Session.findById("wnd[0]").maximize
Session.findById("wnd[0]/tbar[0]/okcd").Text = "/NCJ20N"
Session.findById("wnd[0]").sendVKey 0
Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[0]/shell").pressButton "OPEN"
Session.findById("wnd[1]/usr/ctxtCNPB_W_ADD_OBJ_DYN-PROJ_EXT").Text = ""
Session.findById("wnd[1]/usr/ctxtCNPB_W_ADD_OBJ_DYN-PRPS_EXT").Text = FormularioDados.txtPEP.Value
Session.findById("wnd[1]").sendVKey 0

'Encerrar programa em caso de mensagem de erro (PEP inexistente)
If Session.ActiveWindow.Text = "Erro" Then
    Session.findById("wnd[2]").sendVKey 0
    Session.findById("wnd[1]").Close
    Session.findById("wnd[0]/tbar[0]/btn[15]").press
    MsgBox "PEP informado não existe. Programa interrompido.", vbOKOnly
    Exit Sub
End If

'Por em modo de edição
If Session.ActiveWindow.Text = "Project Builder: exibir subprojeto " & FormularioDados.txtPEP.Value Then
    Session.findById("wnd[0]/tbar[1]/btn[13]").press
    If Session.findById("wnd[0]/sbar").Text = "Não foram bloqueados todos os objetos (ver protocolo de bloqueio)" Then
        Session.findById("wnd[0]/tbar[0]/btn[15]").press
        MsgBox "PEP está aberto por outro usuário. Verificar e tentar novamente mais tarde.", vbOKOnly
        Exit Sub
    End If
Else
    If Session.findById("wnd[0]/sbar").Text = "Não foram bloqueados todos os objetos (ver protocolo de bloqueio)" Then
        Session.findById("wnd[0]/tbar[0]/btn[15]").press
        MsgBox "PEP está aberto por outro usuário. Verificar e tentar novamente mais tarde.", vbOKOnly
        Exit Sub
    End If
End If

'Buscar tarefa da Expedição
Dim nr_linhas_visiveis As Integer
Dim nr_max_linhas As Integer
Dim x As Integer
Dim mat_ok As Boolean
Dim y As Integer
Dim layout As String

mat_ok = False

Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "ACTY_OVW"
nr_max_linhas = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Maximum + 1
nr_linhas_visiveis = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").visiblerowcount
For x = 0 To nr_linhas_visiveis - 1
    If mat_ok = False Then
        If FormularioDados.tipoestrutura.Value <> "ESSW (elétrica)" Then
            If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-VORNR[0," & x & "]").Text = "0900" Or Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-VORNR[0," & x & "]").Text = "0899" Then
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(x).Selected = True
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/btnBUTTON_COMPONENT_OVERVIEW").press
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOMK:2799/tabsTABSTRIP_2700/tabpALLE/ssubSUBSCR_2000:SAPLCOMK:2701/tblSAPLCOMKTCTRL_2701/btn").press
                If Session.findById("wnd[1]/usr/cmbAKT_VERSION").Key <> "Config.global" Then
                    layout = Session.findById("wnd[1]/usr/cmbAKT_VERSION").Key
                    Session.findById("wnd[1]/usr/cmbAKT_VERSION").Key = "Config.global"
                Else
                    layout = "Config.global"
                End If
                Session.findById("wnd[1]/tbar[0]/btn[11]").press
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOMK:2799/tabsTABSTRIP_2700/tabpALLE/ssubSUBSCR_2000:SAPLCOMK:2701/tblSAPLCOMKTCTRL_2701/ctxtRESBD-MATNR[1,0]").Text = FormularioDados.txtMaterial.Value
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOMK:2799/tabsTABSTRIP_2700/tabpALLE/ssubSUBSCR_2000:SAPLCOMK:2701/tblSAPLCOMKTCTRL_2701/txtRESBD-MENGE[3,0]").Text = "-1"
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOMK:2799/tabsTABSTRIP_2700/tabpALLE/ssubSUBSCR_2000:SAPLCOMK:2701/tblSAPLCOMKTCTRL_2701/ctxtRESBD-POSTP[6,0]").Text = "L"
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOMK:2799/tabsTABSTRIP_2700/tabpALLE/ssubSUBSCR_2000:SAPLCOMK:2701/tblSAPLCOMKTCTRL_2701/ctxtRESBD-MFLIC[5,0]").Text = "ZEL"
                If FormularioDados.chkCPC17.Value = True Then
                    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOMK:2799/tabsTABSTRIP_2700/tabpALLE/ssubSUBSCR_2000:SAPLCOMK:2701/tblSAPLCOMKTCTRL_2701/chkRESBD-SANKA[32,0]").Selected = False
                End If
                Session.findById("wnd[0]").sendVKey 0
                If layout <> "Config.global" Then
                    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOMK:2799/tabsTABSTRIP_2700/tabpALLE/ssubSUBSCR_2000:SAPLCOMK:2701/tblSAPLCOMKTCTRL_2701/btn").press
                    Session.findById("wnd[1]/usr/cmbAKT_VERSION").Key = layout
                    Session.findById("wnd[1]/tbar[0]/btn[11]").press
                End If
                mat_ok = True
                Exit For
            End If
        Else
            If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-VORNR[0," & x & "]").Text = "0300" Or Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-VORNR[0," & x & "]").Text = "0299" Then
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(x).Selected = True
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/btnBUTTON_COMPONENT_OVERVIEW").press
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOMK:2799/tabsTABSTRIP_2700/tabpALLE/ssubSUBSCR_2000:SAPLCOMK:2701/tblSAPLCOMKTCTRL_2701/btn").press
                If Session.findById("wnd[1]/usr/cmbAKT_VERSION").Key <> "Config.global" Then
                    layout = Session.findById("wnd[1]/usr/cmbAKT_VERSION").Key
                    Session.findById("wnd[1]/usr/cmbAKT_VERSION").Key = "Config.global"
                Else
                    layout = "Config.global"
                End If
                Session.findById("wnd[1]/tbar[0]/btn[11]").press
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOMK:2799/tabsTABSTRIP_2700/tabpALLE/ssubSUBSCR_2000:SAPLCOMK:2701/tblSAPLCOMKTCTRL_2701/ctxtRESBD-MATNR[1,0]").Text = FormularioDados.txtMaterial.Value
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOMK:2799/tabsTABSTRIP_2700/tabpALLE/ssubSUBSCR_2000:SAPLCOMK:2701/tblSAPLCOMKTCTRL_2701/txtRESBD-MENGE[3,0]").Text = "-1"
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOMK:2799/tabsTABSTRIP_2700/tabpALLE/ssubSUBSCR_2000:SAPLCOMK:2701/tblSAPLCOMKTCTRL_2701/ctxtRESBD-POSTP[6,0]").Text = "L"
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOMK:2799/tabsTABSTRIP_2700/tabpALLE/ssubSUBSCR_2000:SAPLCOMK:2701/tblSAPLCOMKTCTRL_2701/ctxtRESBD-MFLIC[5,0]").Text = "ZEL"
                If FormularioDados.chkCPC17.Value = True Then
                    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOMK:2799/tabsTABSTRIP_2700/tabpALLE/ssubSUBSCR_2000:SAPLCOMK:2701/tblSAPLCOMKTCTRL_2701/chkRESBD-SANKA[32,0]").Selected = False
                End If
                Session.findById("wnd[0]").sendVKey 0
                If layout <> "Config.global" Then
                    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOMK:2799/tabsTABSTRIP_2700/tabpALLE/ssubSUBSCR_2000:SAPLCOMK:2701/tblSAPLCOMKTCTRL_2701/btn").press
                    Session.findById("wnd[1]/usr/cmbAKT_VERSION").Key = layout
                    Session.findById("wnd[1]/tbar[0]/btn[11]").press
                End If
                mat_ok = True
                Exit For
            End If
        End If
    End If
Next
'Seguir procurando na próxima página caso não tenha encontrado as tarefas na primeira
If mat_ok = False Then
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = nr_max_linhas - nr_linhas_visiveis - 3
    For x = 0 To 3
        If x <> 0 Then
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position + 1
        End If
        If mat_ok = False Then
            If FormularioDados.tipoestrutura.Value <> "ESSW (elétrica)" Then
                If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-VORNR[0,0]").Text = "0899" Or Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-VORNR[0,0]").Text = "0900" Then
                    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position).Selected = True
                    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/btnBUTTON_COMPONENT_OVERVIEW").press
                    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOMK:2799/tabsTABSTRIP_2700/tabpALLE/ssubSUBSCR_2000:SAPLCOMK:2701/tblSAPLCOMKTCTRL_2701/btn").press
                    If Session.findById("wnd[1]/usr/cmbAKT_VERSION").Key <> "Config.global" Then
                        layout = Session.findById("wnd[1]/usr/cmbAKT_VERSION").Key
                        Session.findById("wnd[1]/usr/cmbAKT_VERSION").Key = "Config.global"
                    Else
                        layout = "Config.global"
                    End If
                    Session.findById("wnd[1]/tbar[0]/btn[11]").press
                    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOMK:2799/tabsTABSTRIP_2700/tabpALLE/ssubSUBSCR_2000:SAPLCOMK:2701/tblSAPLCOMKTCTRL_2701/ctxtRESBD-MATNR[1,0]").Text = FormularioDados.txtMaterial.Value
                    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOMK:2799/tabsTABSTRIP_2700/tabpALLE/ssubSUBSCR_2000:SAPLCOMK:2701/tblSAPLCOMKTCTRL_2701/txtRESBD-MENGE[3,0]").Text = "-1"
                    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOMK:2799/tabsTABSTRIP_2700/tabpALLE/ssubSUBSCR_2000:SAPLCOMK:2701/tblSAPLCOMKTCTRL_2701/ctxtRESBD-POSTP[6,0]").Text = "L"
                    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOMK:2799/tabsTABSTRIP_2700/tabpALLE/ssubSUBSCR_2000:SAPLCOMK:2701/tblSAPLCOMKTCTRL_2701/ctxtRESBD-MFLIC[5,0]").Text = "ZEL"
                    If FormularioDados.chkCPC17.Value = True Then
                        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOMK:2799/tabsTABSTRIP_2700/tabpALLE/ssubSUBSCR_2000:SAPLCOMK:2701/tblSAPLCOMKTCTRL_2701/chkRESBD-SANKA[32,0]").Selected = False
                    End If
                    Session.findById("wnd[0]").sendVKey 0
                    If layout <> "Config.global" Then
                        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOMK:2799/tabsTABSTRIP_2700/tabpALLE/ssubSUBSCR_2000:SAPLCOMK:2701/tblSAPLCOMKTCTRL_2701/btn").press
                        Session.findById("wnd[1]/usr/cmbAKT_VERSION").Key = layout
                        Session.findById("wnd[1]/tbar[0]/btn[11]").press
                    End If
                    mat_ok = True
                    Exit For
                End If
            Else
                If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-VORNR[0,0]").Text = "0299" Or Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-VORNR[0,0]").Text = "0300" Then
                    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position).Selected = True
                    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/btnBUTTON_COMPONENT_OVERVIEW").press
                    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOMK:2799/tabsTABSTRIP_2700/tabpALLE/ssubSUBSCR_2000:SAPLCOMK:2701/tblSAPLCOMKTCTRL_2701/btn").press
                    If Session.findById("wnd[1]/usr/cmbAKT_VERSION").Key <> "Config.global" Then
                        layout = Session.findById("wnd[1]/usr/cmbAKT_VERSION").Key
                        Session.findById("wnd[1]/usr/cmbAKT_VERSION").Key = "Config.global"
                    Else
                        layout = "Config.global"
                    End If
                    Session.findById("wnd[1]/tbar[0]/btn[11]").press
                    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOMK:2799/tabsTABSTRIP_2700/tabpALLE/ssubSUBSCR_2000:SAPLCOMK:2701/tblSAPLCOMKTCTRL_2701/ctxtRESBD-MATNR[1,0]").Text = FormularioDados.txtMaterial.Value
                    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOMK:2799/tabsTABSTRIP_2700/tabpALLE/ssubSUBSCR_2000:SAPLCOMK:2701/tblSAPLCOMKTCTRL_2701/txtRESBD-MENGE[3,0]").Text = "-1"
                    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOMK:2799/tabsTABSTRIP_2700/tabpALLE/ssubSUBSCR_2000:SAPLCOMK:2701/tblSAPLCOMKTCTRL_2701/ctxtRESBD-POSTP[6,0]").Text = "L"
                    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOMK:2799/tabsTABSTRIP_2700/tabpALLE/ssubSUBSCR_2000:SAPLCOMK:2701/tblSAPLCOMKTCTRL_2701/ctxtRESBD-MFLIC[5,0]").Text = "ZEL"
                    If FormularioDados.chkCPC17.Value = True Then
                        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOMK:2799/tabsTABSTRIP_2700/tabpALLE/ssubSUBSCR_2000:SAPLCOMK:2701/tblSAPLCOMKTCTRL_2701/chkRESBD-SANKA[32,0]").Selected = False
                    End If
                    Session.findById("wnd[0]").sendVKey 0
                    If layout <> "Config.global" Then
                        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOMK:2799/tabsTABSTRIP_2700/tabpALLE/ssubSUBSCR_2000:SAPLCOMK:2701/tblSAPLCOMKTCTRL_2701/btn").press
                        Session.findById("wnd[1]/usr/cmbAKT_VERSION").Key = layout
                        Session.findById("wnd[1]/tbar[0]/btn[11]").press
                    End If
                    mat_ok = True
                    Exit For
                End If
            End If
        End If
    Next
End If



If FormularioDados.txtMateriaavo.Value <> "" Then

mat_ok = False

Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "ACTY_OVW"
nr_max_linhas = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Maximum + 1
nr_linhas_visiveis = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").visiblerowcount
For x = 0 To nr_linhas_visiveis - 1
    If mat_ok = False Then
        If FormularioDados.tipoestrutura.Value <> "ESSW (elétrica)" Then
            If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-VORNR[0," & x & "]").Text = "0905" Then
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(x).Selected = True
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/btnBUTTON_COMPONENT_OVERVIEW").press
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOMK:2799/tabsTABSTRIP_2700/tabpALLE/ssubSUBSCR_2000:SAPLCOMK:2701/tblSAPLCOMKTCTRL_2701/btn").press
                If Session.findById("wnd[1]/usr/cmbAKT_VERSION").Key <> "Config.global" Then
                    layout = Session.findById("wnd[1]/usr/cmbAKT_VERSION").Key
                    Session.findById("wnd[1]/usr/cmbAKT_VERSION").Key = "Config.global"
                Else
                    layout = "Config.global"
                End If
                Session.findById("wnd[1]/tbar[0]/btn[11]").press
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOMK:2799/tabsTABSTRIP_2700/tabpALLE/ssubSUBSCR_2000:SAPLCOMK:2701/tblSAPLCOMKTCTRL_2701/ctxtRESBD-MATNR[1,0]").Text = FormularioDados.txtMateriaavo.Value
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOMK:2799/tabsTABSTRIP_2700/tabpALLE/ssubSUBSCR_2000:SAPLCOMK:2701/tblSAPLCOMKTCTRL_2701/txtRESBD-MENGE[3,0]").Text = "-1"
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOMK:2799/tabsTABSTRIP_2700/tabpALLE/ssubSUBSCR_2000:SAPLCOMK:2701/tblSAPLCOMKTCTRL_2701/ctxtRESBD-POSTP[6,0]").Text = "L"
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOMK:2799/tabsTABSTRIP_2700/tabpALLE/ssubSUBSCR_2000:SAPLCOMK:2701/tblSAPLCOMKTCTRL_2701/ctxtRESBD-MFLIC[5,0]").Text = "ZEL"
                If FormularioDados.chkCPC17.Value = True Then
                    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOMK:2799/tabsTABSTRIP_2700/tabpALLE/ssubSUBSCR_2000:SAPLCOMK:2701/tblSAPLCOMKTCTRL_2701/chkRESBD-SANKA[32,0]").Selected = False
                End If
                Session.findById("wnd[0]").sendVKey 0
                If layout <> "Config.global" Then
                    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOMK:2799/tabsTABSTRIP_2700/tabpALLE/ssubSUBSCR_2000:SAPLCOMK:2701/tblSAPLCOMKTCTRL_2701/btn").press
                    Session.findById("wnd[1]/usr/cmbAKT_VERSION").Key = layout
                    Session.findById("wnd[1]/tbar[0]/btn[11]").press
                End If
                mat_ok = True
                Exit For
            End If
        Else
            If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-VORNR[0," & x & "]").Text = "0300" Or Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-VORNR[0," & x & "]").Text = "0299" Then
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(x).Selected = True
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/btnBUTTON_COMPONENT_OVERVIEW").press
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOMK:2799/tabsTABSTRIP_2700/tabpALLE/ssubSUBSCR_2000:SAPLCOMK:2701/tblSAPLCOMKTCTRL_2701/btn").press
                If Session.findById("wnd[1]/usr/cmbAKT_VERSION").Key <> "Config.global" Then
                    layout = Session.findById("wnd[1]/usr/cmbAKT_VERSION").Key
                    Session.findById("wnd[1]/usr/cmbAKT_VERSION").Key = "Config.global"
                Else
                    layout = "Config.global"
                End If
                Session.findById("wnd[1]/tbar[0]/btn[11]").press
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOMK:2799/tabsTABSTRIP_2700/tabpALLE/ssubSUBSCR_2000:SAPLCOMK:2701/tblSAPLCOMKTCTRL_2701/ctxtRESBD-MATNR[1,0]").Text = FormularioDados.txtMateriaavo.Value
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOMK:2799/tabsTABSTRIP_2700/tabpALLE/ssubSUBSCR_2000:SAPLCOMK:2701/tblSAPLCOMKTCTRL_2701/txtRESBD-MENGE[3,0]").Text = "-1"
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOMK:2799/tabsTABSTRIP_2700/tabpALLE/ssubSUBSCR_2000:SAPLCOMK:2701/tblSAPLCOMKTCTRL_2701/ctxtRESBD-POSTP[6,0]").Text = "L"
                Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOMK:2799/tabsTABSTRIP_2700/tabpALLE/ssubSUBSCR_2000:SAPLCOMK:2701/tblSAPLCOMKTCTRL_2701/ctxtRESBD-MFLIC[5,0]").Text = "ZEL"
                If FormularioDados.chkCPC17.Value = True Then
                    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOMK:2799/tabsTABSTRIP_2700/tabpALLE/ssubSUBSCR_2000:SAPLCOMK:2701/tblSAPLCOMKTCTRL_2701/chkRESBD-SANKA[32,0]").Selected = False
                End If
                Session.findById("wnd[0]").sendVKey 0
                If layout <> "Config.global" Then
                    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOMK:2799/tabsTABSTRIP_2700/tabpALLE/ssubSUBSCR_2000:SAPLCOMK:2701/tblSAPLCOMKTCTRL_2701/btn").press
                    Session.findById("wnd[1]/usr/cmbAKT_VERSION").Key = layout
                    Session.findById("wnd[1]/tbar[0]/btn[11]").press
                End If
                mat_ok = True
                Exit For
            End If
        End If
    End If
Next
'Seguir procurando na próxima página caso não tenha encontrado as tarefas na primeira
If mat_ok = False Then
    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = nr_max_linhas - nr_linhas_visiveis - 3
    For x = 0 To 3
        If x <> 0 Then
            Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position + 1
        End If
        If mat_ok = False Then
            If FormularioDados.tipoestrutura.Value <> "ESSW (elétrica)" Then
                If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-VORNR[0,0]").Text = "0905" Then
                    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position).Selected = True
                    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/btnBUTTON_COMPONENT_OVERVIEW").press
                    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOMK:2799/tabsTABSTRIP_2700/tabpALLE/ssubSUBSCR_2000:SAPLCOMK:2701/tblSAPLCOMKTCTRL_2701/btn").press
                    If Session.findById("wnd[1]/usr/cmbAKT_VERSION").Key <> "Config.global" Then
                        layout = Session.findById("wnd[1]/usr/cmbAKT_VERSION").Key
                        Session.findById("wnd[1]/usr/cmbAKT_VERSION").Key = "Config.global"
                    Else
                        layout = "Config.global"
                    End If
                    Session.findById("wnd[1]/tbar[0]/btn[11]").press
                    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOMK:2799/tabsTABSTRIP_2700/tabpALLE/ssubSUBSCR_2000:SAPLCOMK:2701/tblSAPLCOMKTCTRL_2701/ctxtRESBD-MATNR[1,0]").Text = FormularioDados.txtMateriaavo.Value
                    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOMK:2799/tabsTABSTRIP_2700/tabpALLE/ssubSUBSCR_2000:SAPLCOMK:2701/tblSAPLCOMKTCTRL_2701/txtRESBD-MENGE[3,0]").Text = "-1"
                    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOMK:2799/tabsTABSTRIP_2700/tabpALLE/ssubSUBSCR_2000:SAPLCOMK:2701/tblSAPLCOMKTCTRL_2701/ctxtRESBD-POSTP[6,0]").Text = "L"
                    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOMK:2799/tabsTABSTRIP_2700/tabpALLE/ssubSUBSCR_2000:SAPLCOMK:2701/tblSAPLCOMKTCTRL_2701/ctxtRESBD-MFLIC[5,0]").Text = "ZEL"
                    If FormularioDados.chkCPC17.Value = True Then
                        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOMK:2799/tabsTABSTRIP_2700/tabpALLE/ssubSUBSCR_2000:SAPLCOMK:2701/tblSAPLCOMKTCTRL_2701/chkRESBD-SANKA[32,0]").Selected = False
                    End If
                    Session.findById("wnd[0]").sendVKey 0
                    If layout <> "Config.global" Then
                        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOMK:2799/tabsTABSTRIP_2700/tabpALLE/ssubSUBSCR_2000:SAPLCOMK:2701/tblSAPLCOMKTCTRL_2701/btn").press
                        Session.findById("wnd[1]/usr/cmbAKT_VERSION").Key = layout
                        Session.findById("wnd[1]/tbar[0]/btn[11]").press
                    End If
                    mat_ok = True
                    Exit For
                End If
            Else
                If Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-VORNR[0,0]").Text = "0299" Or Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030/txtAFVGD-VORNR[0,0]").Text = "0300" Then
                    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").getAbsoluteRow(Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/tblSAPLCOVGTCTRL_2030").verticalScrollbar.Position).Selected = True
                    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpARBL/ssubSUBSCR_2000:SAPLCOVG:2030/btnBUTTON_COMPONENT_OVERVIEW").press
                    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOMK:2799/tabsTABSTRIP_2700/tabpALLE/ssubSUBSCR_2000:SAPLCOMK:2701/tblSAPLCOMKTCTRL_2701/btn").press
                    If Session.findById("wnd[1]/usr/cmbAKT_VERSION").Key <> "Config.global" Then
                        layout = Session.findById("wnd[1]/usr/cmbAKT_VERSION").Key
                        Session.findById("wnd[1]/usr/cmbAKT_VERSION").Key = "Config.global"
                    Else
                        layout = "Config.global"
                    End If
                    Session.findById("wnd[1]/tbar[0]/btn[11]").press
                    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOMK:2799/tabsTABSTRIP_2700/tabpALLE/ssubSUBSCR_2000:SAPLCOMK:2701/tblSAPLCOMKTCTRL_2701/ctxtRESBD-MATNR[1,0]").Text = FormularioDados.txtMateriaavo.Value
                    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOMK:2799/tabsTABSTRIP_2700/tabpALLE/ssubSUBSCR_2000:SAPLCOMK:2701/tblSAPLCOMKTCTRL_2701/txtRESBD-MENGE[3,0]").Text = "-1"
                    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOMK:2799/tabsTABSTRIP_2700/tabpALLE/ssubSUBSCR_2000:SAPLCOMK:2701/tblSAPLCOMKTCTRL_2701/ctxtRESBD-POSTP[6,0]").Text = "L"
                    Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOMK:2799/tabsTABSTRIP_2700/tabpALLE/ssubSUBSCR_2000:SAPLCOMK:2701/tblSAPLCOMKTCTRL_2701/ctxtRESBD-MFLIC[5,0]").Text = "ZEL"
                    If FormularioDados.chkCPC17.Value = True Then
                        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOMK:2799/tabsTABSTRIP_2700/tabpALLE/ssubSUBSCR_2000:SAPLCOMK:2701/tblSAPLCOMKTCTRL_2701/chkRESBD-SANKA[32,0]").Selected = False
                    End If
                    Session.findById("wnd[0]").sendVKey 0
                    If layout <> "Config.global" Then
                        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOMK:2799/tabsTABSTRIP_2700/tabpALLE/ssubSUBSCR_2000:SAPLCOMK:2701/tblSAPLCOMKTCTRL_2701/btn").press
                        Session.findById("wnd[1]/usr/cmbAKT_VERSION").Key = layout
                        Session.findById("wnd[1]/tbar[0]/btn[11]").press
                    End If
                    mat_ok = True
                    Exit For
                End If
            End If
        End If
    Next
End If
End If


If FormularioDados.chkSolar.Value = True Then
    If matfat_ok = True Or mat_ok = True Then
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_DETAIL/shellcont/shell").pressButton "ACTY_DET"
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCONW:1001/tabsTABSTRIP_1000/tabpUSER").Select
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCONW:1001/tabsTABSTRIP_1000/tabpUSER/ssubSUBSCR_1000:SAPLCONW:1320/ctxtAFVGD-USR03").Text = "SOL"
        Session.findById("wnd[0]").sendVKey 0
    End If
End If


'Preencher quando tem 2 negativos
If FormularioDados.Agrupamento.Value = True And FormularioDados.tipoestrutura.Value = "Container Solar" And FormularioDados.proBTI.Value = True Then

'Diagrama de produçãoBTI
Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "NETW_OVW"
    If FormularioDados.SemEng.Value = True Then
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtNETW_OVW-AUFNR[0,3]").SetFocus
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtNETW_OVW-AUFNR[0,3]").caretPosition = 9
        Else
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtNETW_OVW-AUFNR[0,5]").SetFocus
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtNETW_OVW-AUFNR[0,5]").caretPosition = 9
    End If
Session.findById("wnd[0]").sendVKey 2
Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabp+CU3").Select
Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabp+CU3/ssubSUBSCR_2100:SAPLCOKO:2998/ssubCUSTSCR1:SAPLXCN1:0900/chkWA_ZTBFI_188-ID_AGRUPADOR").SetFocus
Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabp+CU3/ssubSUBSCR_2100:SAPLCOKO:2998/ssubCUSTSCR1:SAPLXCN1:0900/chkWA_ZTBFI_188-ID_AGRUPADOR").Selected = True

'Diagrama Elétrica WIJ
Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").topNode = "000001"
Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "NETW_OVW"
    If FormularioDados.SemEng.Value = True Then
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtNETW_OVW-AUFNR[0,2]").SetFocus
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtNETW_OVW-AUFNR[0,2]").caretPosition = 9
        Else
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtNETW_OVW-AUFNR[0,4]").SetFocus
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtNETW_OVW-AUFNR[0,4]").caretPosition = 5
    End If
Session.findById("wnd[0]").sendVKey 2
Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabp+CU3").Select
Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabp+CU3/ssubSUBSCR_2100:SAPLCOKO:2998/ssubCUSTSCR1:SAPLXCN1:0900/chkWA_ZTBFI_188-ID_AGRUPADOR").SetFocus
Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabp+CU3/ssubSUBSCR_2100:SAPLCOKO:2998/ssubCUSTSCR1:SAPLXCN1:0900/chkWA_ZTBFI_188-ID_AGRUPADOR").Selected = True
   
 
 'Diagrama acessório WIJ
Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").topNode = "000001"
Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "NETW_OVW"
    If FormularioDados.SemEng.Value = True Then
        DIGELEC = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtNETW_OVW-AUFNR[0,2]").Text
        DIGPROBTI = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtNETW_OVW-AUFNR[0,3]").Text
        Else
        DIGELEC = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtNETW_OVW-AUFNR[0,4]").Text
        DIGPROBTI = Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtNETW_OVW-AUFNR[0,5]").Text
    End If
    If FormularioDados.SemEng.Value = True Then
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtNETW_OVW-AUFNR[0,1]").SetFocus
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtNETW_OVW-AUFNR[0,1]").caretPosition = 7
        Else
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtNETW_OVW-AUFNR[0,3]").SetFocus
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtNETW_OVW-AUFNR[0,3]").caretPosition = 7
    End If
Session.findById("wnd[0]").sendVKey 2
Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabp+CU3").Select
Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabp+CU3/ssubSUBSCR_2100:SAPLCOKO:2998/ssubCUSTSCR1:SAPLXCN1:0900/ctxtWA_ZTBFI_188-CD_AGRUPAMENTO").Text = DIGELEC
Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabp+CU3/ssubSUBSCR_2100:SAPLCOKO:2998/ssubCUSTSCR1:SAPLXCN1:0900/ctxtWA_ZTBFI_188-CD_AGRUPAMENTO").SetFocus
Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabp+CU3/ssubSUBSCR_2100:SAPLCOKO:2998/ssubCUSTSCR1:SAPLXCN1:0900/ctxtWA_ZTBFI_188-CD_AGRUPAMENTO").caretPosition = 9
Session.findById("wnd[0]").sendVKey 2

'Diagrama MEC WIJ
Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").topNode = "000001"
Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "NETW_OVW"
    If FormularioDados.SemEng.Value = True Then
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtNETW_OVW-AUFNR[0,0]").SetFocus
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtNETW_OVW-AUFNR[0,0]").caretPosition = 7
        Else
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtNETW_OVW-AUFNR[0,2]").SetFocus
        Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtNETW_OVW-AUFNR[0,2]").caretPosition = 7
    End If
Session.findById("wnd[0]").sendVKey 2
Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabp+CU3").Select
Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabp+CU3/ssubSUBSCR_2100:SAPLCOKO:2998/ssubCUSTSCR1:SAPLXCN1:0900/ctxtWA_ZTBFI_188-CD_AGRUPAMENTO").Text = DIGELEC
Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabp+CU3/ssubSUBSCR_2100:SAPLCOKO:2998/ssubCUSTSCR1:SAPLXCN1:0900/ctxtWA_ZTBFI_188-CD_AGRUPAMENTO").SetFocus
Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabp+CU3/ssubSUBSCR_2100:SAPLCOKO:2998/ssubCUSTSCR1:SAPLXCN1:0900/ctxtWA_ZTBFI_188-CD_AGRUPAMENTO").caretPosition = 9
Session.findById("wnd[0]").sendVKey 2
End If

'Reprogramar novamente
Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000002"
Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_DETAIL/shellcont/shell").pressButton "WBSE_DET"
Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCJWB:3999/tabsTABCJWB/tabpTERM").Select
Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCJWB:3999/tabsTABCJWB/tabpTERM/ssubSUBSCR1:SAPLCJTR:3500/ctxtPRTE-PSTRT").Text = FormularioDados.txtDataInicio.Value
Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCJWB:3999/tabsTABCJWB/tabpTERM/ssubSUBSCR1:SAPLCJTR:3500/txtPRTE-PDAUR").Text = ""
Session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCJWB:3999/tabsTABCJWB/tabpTERM/ssubSUBSCR1:SAPLCJTR:3500/ctxtPRTE-PENDE").Text = ""
Session.findById("wnd[0]").sendVKey 0
Session.findById("wnd[0]").sendVKey 33


'Salvar
Session.findById("wnd[0]/tbar[0]/btn[11]").press

'Aviso orçamento de custos CPC17
If Session.ActiveWindow.Text = "Ctrl.disponibilidade" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press

'Erro na determinação dos custos
If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Programação" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
If Session.ActiveWindow.Text = "Determ.custos" Then Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press


'Aviso Bloco K
If Session.ActiveWindow.Text = "Baixa existente em período anterior" Then Session.findById("wnd[1]/usr/btnBUTTON_1").press

'Sair da CJ20N
Session.findById("wnd[0]/tbar[0]/btn[15]").press

'Application.DisplayAlerts = True


End Sub
```

#### `inicio_programa` (Sub)

```vb
Sub inicio_programa()

Application.DisplayAlerts = False

'Checagem de inconsistência/falta de dados
If FormularioDados.tipoestrutura.Value = "" Then
    MsgBox "Tipo de estrutura não selecionado. Favor verificar.", vbOKOnly
    Exit Sub
End If

If FormularioDados.nrmodulos.Value = "" Then
    MsgBox "Nº de módulos não selecionado. Favor verificar.", vbOKOnly
    Exit Sub
End If

If FormularioDados.planpin.Value = "" Then
    MsgBox "Plano de Pintura não selecionado. Favor verificar.", vbOKOnly
    Exit Sub
End If

'Comprimento = ""
If FormularioDados.modulo1.Value = "" And FormularioDados.tipoestrutura.Value <> "Container Solar" And FormularioDados.tipoestrutura.Value <> "Skid (mecânica)" And FormularioDados.tipoestrutura.Value <> "Skid (com elétrica)" And FormularioDados.tipoestrutura.Value <> "Pilotis" And FormularioDados.tipoestrutura.Value <> "ESSW (mecânica)" And FormularioDados.tipoestrutura.Value <> "ESSW (elétrica)" And FormularioDados.tipoestrutura.Value <> "Serviço Engenharia" Then
    MsgBox "Comprimento do Módulo 1 não informado. Favor verificar.", vbOKOnly
    Exit Sub
End If

If FormularioDados.modulo2.Value = "" And FormularioDados.tipoestrutura.Value <> "Container Solar" And FormularioDados.tipoestrutura.Value <> "Skid (mecânica)" And FormularioDados.tipoestrutura.Value <> "Skid (com elétrica)" And FormularioDados.tipoestrutura.Value <> "Pilotis" And FormularioDados.tipoestrutura.Value <> "ESSW (mecânica)" And FormularioDados.tipoestrutura.Value <> "ESSW (elétrica)" And FormularioDados.nrmodulos.Value <> "1 Módulo" And FormularioDados.tipoestrutura.Value <> "Serviço Engenharia" Then
    MsgBox "Comprimento do Módulo 2 não informado. Favor verificar.", vbOKOnly
    Exit Sub
End If

If FormularioDados.modulo3.Value = "" And FormularioDados.tipoestrutura.Value <> "Container Solar" And FormularioDados.tipoestrutura.Value <> "Skid (mecânica)" And FormularioDados.tipoestrutura.Value <> "Skid (com elétrica)" And FormularioDados.tipoestrutura.Value <> "Pilotis" And FormularioDados.tipoestrutura.Value <> "ESSW (mecânica)" And FormularioDados.tipoestrutura.Value <> "ESSW (elétrica)" And FormularioDados.nrmodulos.Value <> "1 Módulo" And FormularioDados.nrmodulos.Value <> "2 Módulos" And FormularioDados.tipoestrutura.Value <> "Serviço Engenharia" Then
    MsgBox "Comprimento do Módulo 3 não informado. Favor verificar.", vbOKOnly
    Exit Sub
End If

If FormularioDados.modulo4.Value = "" And FormularioDados.tipoestrutura.Value <> "Container Solar" And FormularioDados.tipoestrutura.Value <> "Skid (mecânica)" And FormularioDados.tipoestrutura.Value <> "Skid (com elétrica)" And FormularioDados.tipoestrutura.Value <> "Pilotis" And FormularioDados.tipoestrutura.Value <> "ESSW (mecânica)" And FormularioDados.tipoestrutura.Value <> "ESSW (elétrica)" And FormularioDados.nrmodulos.Value <> "1 Módulo" And FormularioDados.nrmodulos.Value <> "2 Módulos" And FormularioDados.nrmodulos.Value <> "3 Módulos" And FormularioDados.tipoestrutura.Value <> "Serviço Engenharia" Then
    MsgBox "Comprimento do Módulo 4 não informado. Favor verificar.", vbOKOnly
    Exit Sub
End If

If FormularioDados.modulo4.Value = "" And FormularioDados.tipoestrutura.Value <> "Container Solar" And FormularioDados.tipoestrutura.Value <> "Skid (mecânica)" And FormularioDados.tipoestrutura.Value <> "Skid (com elétrica)" And FormularioDados.tipoestrutura.Value <> "Pilotis" And FormularioDados.tipoestrutura.Value <> "ESSW (mecânica)" And FormularioDados.tipoestrutura.Value <> "ESSW (elétrica)" And FormularioDados.nrmodulos.Value <> "1 Módulo" And FormularioDados.nrmodulos.Value <> "2 Módulos" And FormularioDados.nrmodulos.Value <> "3 Módulos" And FormularioDados.nrmodulos.Value <> "4 Módulos" And FormularioDados.tipoestrutura.Value <> "Serviço Engenharia" Then
    MsgBox "Comprimento do Módulo 5 não informado. Favor verificar.", vbOKOnly
    Exit Sub
End If
If FormularioDados.modulo4.Value = "" And FormularioDados.tipoestrutura.Value <> "Container Solar" And FormularioDados.tipoestrutura.Value <> "Skid (mecânica)" And FormularioDados.tipoestrutura.Value <> "Skid (com elétrica)" And FormularioDados.tipoestrutura.Value <> "Pilotis" And FormularioDados.tipoestrutura.Value <> "ESSW (mecânica)" And FormularioDados.tipoestrutura.Value <> "ESSW (elétrica)" And FormularioDados.nrmodulos.Value <> "1 Módulo" And FormularioDados.nrmodulos.Value <> "2 Módulos" And FormularioDados.nrmodulos.Value <> "3 Módulos" And FormularioDados.nrmodulos.Value <> "4 Módulos" And FormularioDados.nrmodulos.Value <> "5 Módulos" And FormularioDados.tipoestrutura.Value <> "Serviço Engenharia" Then
    MsgBox "Comprimento do Módulo 6 não informado. Favor verificar.", vbOKOnly
    Exit Sub
End If

If FormularioDados.modulo4.Value = "" And FormularioDados.tipoestrutura.Value <> "Container Solar" And FormularioDados.tipoestrutura.Value <> "Skid (mecânica)" And FormularioDados.tipoestrutura.Value <> "Skid (com elétrica)" And FormularioDados.tipoestrutura.Value <> "Pilotis" And FormularioDados.tipoestrutura.Value <> "ESSW (mecânica)" And FormularioDados.tipoestrutura.Value <> "ESSW (elétrica)" And FormularioDados.nrmodulos.Value <> "1 Módulo" And FormularioDados.nrmodulos.Value <> "2 Módulos" And FormularioDados.nrmodulos.Value <> "3 Módulos" And FormularioDados.nrmodulos.Value <> "4 Módulos" And FormularioDados.nrmodulos.Value <> "5 Módulos" And FormularioDados.nrmodulos.Value <> "6 Módulos" And FormularioDados.tipoestrutura.Value <> "Serviço Engenharia" Then
    MsgBox "Comprimento do Módulo 7 não informado. Favor verificar.", vbOKOnly
    Exit Sub
End If

If FormularioDados.modulo4.Value = "" And FormularioDados.tipoestrutura.Value <> "Container Solar" And FormularioDados.tipoestrutura.Value <> "Skid (mecânica)" And FormularioDados.tipoestrutura.Value <> "Skid (com elétrica)" And FormularioDados.tipoestrutura.Value <> "Pilotis" And FormularioDados.tipoestrutura.Value <> "ESSW (mecânica)" And FormularioDados.tipoestrutura.Value <> "ESSW (elétrica)" And FormularioDados.nrmodulos.Value <> "1 Módulo" And FormularioDados.nrmodulos.Value <> "2 Módulos" And FormularioDados.nrmodulos.Value <> "3 Módulos" And FormularioDados.nrmodulos.Value <> "4 Módulos" And FormularioDados.nrmodulos.Value <> "5 Módulos" And FormularioDados.nrmodulos.Value <> "6 Módulos" And FormularioDados.nrmodulos.Value <> "7 Módulos" And FormularioDados.tipoestrutura.Value <> "Serviço Engenharia" Then
    MsgBox "Comprimento do Módulo 8 não informado. Favor verificar.", vbOKOnly
    Exit Sub
End If

'Comprimento = 0
If FormularioDados.modulo1.Value = 0 And FormularioDados.tipoestrutura.Value <> "Container Solar" And FormularioDados.tipoestrutura.Value <> "Skid (mecânica)" And FormularioDados.tipoestrutura.Value <> "Skid (com elétrica)" And FormularioDados.tipoestrutura.Value <> "Pilotis" And FormularioDados.tipoestrutura.Value <> "ESSW (mecânica)" And FormularioDados.tipoestrutura.Value <> "ESSW (elétrica)" And FormularioDados.tipoestrutura.Value <> "Serviço Engenharia" Then
    MsgBox "Comprimento do Módulo 1 = não informado. Favor verificar.", vbOKOnly
    Exit Sub
End If

If FormularioDados.modulo2.Value = 0 And FormularioDados.tipoestrutura.Value <> "Container Solar" And FormularioDados.tipoestrutura.Value <> "Skid (mecânica)" And FormularioDados.tipoestrutura.Value <> "Skid (com elétrica)" And FormularioDados.tipoestrutura.Value <> "Pilotis" And FormularioDados.tipoestrutura.Value <> "ESSW (mecânica)" And FormularioDados.tipoestrutura.Value <> "ESSW (elétrica)" And FormularioDados.nrmodulos.Value <> "1 Módulo" And FormularioDados.tipoestrutura.Value <> "Serviço Engenharia" Then
    MsgBox "Comprimento do Módulo 2 = não informado. Favor verificar.", vbOKOnly
    Exit Sub
End If

If FormularioDados.modulo3.Value = 0 And FormularioDados.tipoestrutura.Value <> "Container Solar" And FormularioDados.tipoestrutura.Value <> "Skid (mecânica)" And FormularioDados.tipoestrutura.Value <> "Skid (com elétrica)" And FormularioDados.tipoestrutura.Value <> "Pilotis" And FormularioDados.tipoestrutura.Value <> "ESSW (mecânica)" And FormularioDados.tipoestrutura.Value <> "ESSW (elétrica)" And FormularioDados.nrmodulos.Value <> "1 Módulo" And FormularioDados.nrmodulos.Value <> "2 Módulos" And FormularioDados.tipoestrutura.Value <> "Serviço Engenharia" Then
    MsgBox "Comprimento do Módulo 3 não informado. Favor verificar.", vbOKOnly
    Exit Sub
End If

If FormularioDados.modulo4.Value = 0 And FormularioDados.tipoestrutura.Value <> "Container Solar" And FormularioDados.tipoestrutura.Value <> "Skid (mecânica)" And FormularioDados.tipoestrutura.Value <> "Skid (com elétrica)" And FormularioDados.tipoestrutura.Value <> "Pilotis" And FormularioDados.tipoestrutura.Value <> "ESSW (mecânica)" And FormularioDados.tipoestrutura.Value <> "ESSW (elétrica)" And FormularioDados.nrmodulos.Value <> "1 Módulo" And FormularioDados.nrmodulos.Value <> "2 Módulos" And FormularioDados.nrmodulos.Value <> "3 Módulos" And FormularioDados.tipoestrutura.Value <> "Serviço Engenharia" Then
    MsgBox "Comprimento do Módulo 4 não informado. Favor verificar.", vbOKOnly
    Exit Sub
End If

If FormularioDados.modulo4.Value = 0 And FormularioDados.tipoestrutura.Value <> "Container Solar" And FormularioDados.tipoestrutura.Value <> "Skid (mecânica)" And FormularioDados.tipoestrutura.Value <> "Skid (com elétrica)" And FormularioDados.tipoestrutura.Value <> "Pilotis" And FormularioDados.tipoestrutura.Value <> "ESSW (mecânica)" And FormularioDados.tipoestrutura.Value <> "ESSW (elétrica)" And FormularioDados.nrmodulos.Value <> "1 Módulo" And FormularioDados.nrmodulos.Value <> "2 Módulos" And FormularioDados.nrmodulos.Value <> "3 Módulos" And FormularioDados.nrmodulos.Value <> "4 Módulos" And FormularioDados.tipoestrutura.Value <> "Serviço Engenharia" Then
    MsgBox "Comprimento do Módulo 5 não informado. Favor verificar.", vbOKOnly
    Exit Sub
End If

If FormularioDados.modulo4.Value = 0 And FormularioDados.tipoestrutura.Value <> "Container Solar" And FormularioDados.tipoestrutura.Value <> "Skid (mecânica)" And FormularioDados.tipoestrutura.Value <> "Skid (com elétrica)" And FormularioDados.tipoestrutura.Value <> "Pilotis" And FormularioDados.tipoestrutura.Value <> "ESSW (mecânica)" And FormularioDados.tipoestrutura.Value <> "ESSW (elétrica)" And FormularioDados.nrmodulos.Value <> "1 Módulo" And FormularioDados.nrmodulos.Value <> "2 Módulos" And FormularioDados.nrmodulos.Value <> "3 Módulos" And FormularioDados.nrmodulos.Value <> "4 Módulos" And FormularioDados.nrmodulos.Value <> "5 Módulos" And FormularioDados.tipoestrutura.Value <> "Serviço Engenharia" Then
    MsgBox "Comprimento do Módulo 6 não informado. Favor verificar.", vbOKOnly
    Exit Sub
End If

If FormularioDados.modulo4.Value = 0 And FormularioDados.tipoestrutura.Value <> "Container Solar" And FormularioDados.tipoestrutura.Value <> "Skid (mecânica)" And FormularioDados.tipoestrutura.Value <> "Skid (com elétrica)" And FormularioDados.tipoestrutura.Value <> "Pilotis" And FormularioDados.tipoestrutura.Value <> "ESSW (mecânica)" And FormularioDados.tipoestrutura.Value <> "ESSW (elétrica)" And FormularioDados.nrmodulos.Value <> "1 Módulo" And FormularioDados.nrmodulos.Value <> "2 Módulos" And FormularioDados.nrmodulos.Value <> "3 Módulos" And FormularioDados.nrmodulos.Value <> "4 Módulos" And FormularioDados.nrmodulos.Value <> "5 Módulos" And FormularioDados.nrmodulos.Value <> "6 Módulos" And FormularioDados.tipoestrutura.Value <> "Serviço Engenharia" Then
    MsgBox "Comprimento do Módulo 7 não informado. Favor verificar.", vbOKOnly
    Exit Sub
End If

If FormularioDados.modulo4.Value = 0 And FormularioDados.tipoestrutura.Value <> "Container Solar" And FormularioDados.tipoestrutura.Value <> "Skid (mecânica)" And FormularioDados.tipoestrutura.Value <> "Skid (com elétrica)" And FormularioDados.tipoestrutura.Value <> "Pilotis" And FormularioDados.tipoestrutura.Value <> "ESSW (mecânica)" And FormularioDados.tipoestrutura.Value <> "ESSW (elétrica)" And FormularioDados.nrmodulos.Value <> "1 Módulo" And FormularioDados.nrmodulos.Value <> "2 Módulos" And FormularioDados.nrmodulos.Value <> "3 Módulos" And FormularioDados.nrmodulos.Value <> "4 Módulos" And FormularioDados.nrmodulos.Value <> "5 Módulos" And FormularioDados.nrmodulos.Value <> "6 Módulos" And FormularioDados.nrmodulos.Value <> "7 Módulos" And FormularioDados.tipoestrutura.Value <> "Serviço Engenharia" Then
    MsgBox "Comprimento do Módulo 8 não informado. Favor verificar.", vbOKOnly
    Exit Sub
End If

'largura = ""
If FormularioDados.largmodulo1.Value = "" And FormularioDados.tipoestrutura.Value <> "Container Solar" And FormularioDados.tipoestrutura.Value <> "Skid (mecânica)" And FormularioDados.tipoestrutura.Value <> "Skid (com elétrica)" And FormularioDados.tipoestrutura.Value <> "Pilotis" And FormularioDados.tipoestrutura.Value <> "ESSW (mecânica)" And FormularioDados.tipoestrutura.Value <> "ESSW (elétrica)" And FormularioDados.tipoestrutura.Value <> "Serviço Engenharia" Then
    MsgBox "Largura do Módulo 1 não informada. Favor verificar.", vbOKOnly
    Exit Sub
End If

If FormularioDados.largmodulo2.Value = "" And FormularioDados.tipoestrutura.Value <> "Container Solar" And FormularioDados.tipoestrutura.Value <> "Skid (mecânica)" And FormularioDados.tipoestrutura.Value <> "Skid (com elétrica)" And FormularioDados.tipoestrutura.Value <> "Pilotis" And FormularioDados.tipoestrutura.Value <> "ESSW (mecânica)" And FormularioDados.tipoestrutura.Value <> "ESSW (elétrica)" And FormularioDados.nrmodulos.Value <> "1 Módulo" And FormularioDados.tipoestrutura.Value <> "Serviço Engenharia" Then
    MsgBox "Largura do Módulo 2 não informada. Favor verificar.", vbOKOnly
    Exit Sub
End If

If FormularioDados.largmodulo3.Value = "" And FormularioDados.tipoestrutura.Value <> "Container Solar" And FormularioDados.tipoestrutura.Value <> "Skid (mecânica)" And FormularioDados.tipoestrutura.Value <> "Skid (com elétrica)" And FormularioDados.tipoestrutura.Value <> "Pilotis" And FormularioDados.tipoestrutura.Value <> "ESSW (mecânica)" And FormularioDados.tipoestrutura.Value <> "ESSW (elétrica)" And FormularioDados.nrmodulos.Value <> "1 Módulo" And FormularioDados.nrmodulos.Value <> "2 Módulos" And FormularioDados.tipoestrutura.Value <> "Serviço Engenharia" Then
    MsgBox "Largura do Módulo 3 não informada. Favor verificar.", vbOKOnly
    Exit Sub
End If

If FormularioDados.largmodulo4.Value = "" And FormularioDados.tipoestrutura.Value <> "Container Solar" And FormularioDados.tipoestrutura.Value <> "Skid (mecânica)" And FormularioDados.tipoestrutura.Value <> "Skid (com elétrica)" And FormularioDados.tipoestrutura.Value <> "Pilotis" And FormularioDados.tipoestrutura.Value <> "ESSW (mecânica)" And FormularioDados.tipoestrutura.Value <> "ESSW (elétrica)" And FormularioDados.nrmodulos.Value <> "1 Módulo" And FormularioDados.nrmodulos.Value <> "2 Módulos" And FormularioDados.nrmodulos.Value <> "3 Módulos" And FormularioDados.tipoestrutura.Value <> "Serviço Engenharia" Then
    MsgBox "Largura do Módulo 4 não informada. Favor verificar.", vbOKOnly
    Exit Sub
End If

If FormularioDados.largmodulo4.Value = "" And FormularioDados.tipoestrutura.Value <> "Container Solar" And FormularioDados.tipoestrutura.Value <> "Skid (mecânica)" And FormularioDados.tipoestrutura.Value <> "Skid (com elétrica)" And FormularioDados.tipoestrutura.Value <> "Pilotis" And FormularioDados.tipoestrutura.Value <> "ESSW (mecânica)" And FormularioDados.tipoestrutura.Value <> "ESSW (elétrica)" And FormularioDados.nrmodulos.Value <> "1 Módulo" And FormularioDados.nrmodulos.Value <> "2 Módulos" And FormularioDados.nrmodulos.Value <> "3 Módulos" And FormularioDados.nrmodulos.Value <> "4 Módulos" And FormularioDados.tipoestrutura.Value <> "Serviço Engenharia" Then
    MsgBox "Largura do Módulo 5 não informada. Favor verificar.", vbOKOnly
    Exit Sub
End If

If FormularioDados.largmodulo4.Value = "" And FormularioDados.tipoestrutura.Value <> "Container Solar" And FormularioDados.tipoestrutura.Value <> "Skid (mecânica)" And FormularioDados.tipoestrutura.Value <> "Skid (com elétrica)" And FormularioDados.tipoestrutura.Value <> "Pilotis" And FormularioDados.tipoestrutura.Value <> "ESSW (mecânica)" And FormularioDados.tipoestrutura.Value <> "ESSW (elétrica)" And FormularioDados.nrmodulos.Value <> "1 Módulo" And FormularioDados.nrmodulos.Value <> "2 Módulos" And FormularioDados.nrmodulos.Value <> "3 Módulos" And FormularioDados.nrmodulos.Value <> "4 Módulos" And FormularioDados.nrmodulos.Value <> "5 Módulos" And FormularioDados.tipoestrutura.Value <> "Serviço Engenharia" Then
    MsgBox "Largura do Módulo 6 não informada. Favor verificar.", vbOKOnly
    Exit Sub
End If

If FormularioDados.largmodulo4.Value = "" And FormularioDados.tipoestrutura.Value <> "Container Solar" And FormularioDados.tipoestrutura.Value <> "Skid (mecânica)" And FormularioDados.tipoestrutura.Value <> "Skid (com elétrica)" And FormularioDados.tipoestrutura.Value <> "Pilotis" And FormularioDados.tipoestrutura.Value <> "ESSW (mecânica)" And FormularioDados.tipoestrutura.Value <> "ESSW (elétrica)" And FormularioDados.nrmodulos.Value <> "1 Módulo" And FormularioDados.nrmodulos.Value <> "2 Módulos" And FormularioDados.nrmodulos.Value <> "3 Módulos" And FormularioDados.nrmodulos.Value <> "4 Módulos" And FormularioDados.nrmodulos.Value <> "5 Módulos" And FormularioDados.nrmodulos.Value <> "6 Módulos" And FormularioDados.tipoestrutura.Value <> "Serviço Engenharia" Then
    MsgBox "Largura do Módulo 7 não informada. Favor verificar.", vbOKOnly
    Exit Sub
End If

If FormularioDados.largmodulo4.Value = "" And FormularioDados.tipoestrutura.Value <> "Container Solar" And FormularioDados.tipoestrutura.Value <> "Skid (mecânica)" And FormularioDados.tipoestrutura.Value <> "Skid (com elétrica)" And FormularioDados.tipoestrutura.Value <> "Pilotis" And FormularioDados.tipoestrutura.Value <> "ESSW (mecânica)" And FormularioDados.tipoestrutura.Value <> "ESSW (elétrica)" And FormularioDados.nrmodulos.Value <> "1 Módulo" And FormularioDados.nrmodulos.Value <> "2 Módulos" And FormularioDados.nrmodulos.Value <> "3 Módulos" And FormularioDados.nrmodulos.Value <> "4 Módulos" And FormularioDados.nrmodulos.Value <> "5 Módulos" And FormularioDados.nrmodulos.Value <> "6 Módulos" And FormularioDados.nrmodulos.Value <> "7 Módulos" And FormularioDados.tipoestrutura.Value <> "Serviço Engenharia" Then
    MsgBox "Largura do Módulo 8 não informada. Favor verificar.", vbOKOnly
    Exit Sub
End If

'Largura = 0
If FormularioDados.largmodulo1.Value = 0 And FormularioDados.tipoestrutura.Value <> "Container Solar" And FormularioDados.tipoestrutura.Value <> "Skid (mecânica)" And FormularioDados.tipoestrutura.Value <> "Skid (com elétrica)" And FormularioDados.tipoestrutura.Value <> "Pilotis" And FormularioDados.tipoestrutura.Value <> "ESSW (mecânica)" And FormularioDados.tipoestrutura.Value <> "ESSW (elétrica)" And FormularioDados.tipoestrutura.Value <> "Serviço Engenharia" Then
    MsgBox "Largura do Módulo 1 = não informada. Favor verificar.", vbOKOnly
    Exit Sub
End If

If FormularioDados.largmodulo2.Value = 0 And FormularioDados.tipoestrutura.Value <> "Container Solar" And FormularioDados.tipoestrutura.Value <> "Skid (mecânica)" And FormularioDados.tipoestrutura.Value <> "Skid (com elétrica)" And FormularioDados.tipoestrutura.Value <> "Pilotis" And FormularioDados.tipoestrutura.Value <> "ESSW (mecânica)" And FormularioDados.tipoestrutura.Value <> "ESSW (elétrica)" And FormularioDados.nrmodulos.Value <> "1 Módulo" And FormularioDados.tipoestrutura.Value <> "Serviço Engenharia" Then
    MsgBox "Largura do Módulo 2 = não informada. Favor verificar.", vbOKOnly
    Exit Sub
End If

If FormularioDados.largmodulo3.Value = 0 And FormularioDados.tipoestrutura.Value <> "Container Solar" And FormularioDados.tipoestrutura.Value <> "Skid (mecânica)" And FormularioDados.tipoestrutura.Value <> "Skid (com elétrica)" And FormularioDados.tipoestrutura.Value <> "Pilotis" And FormularioDados.tipoestrutura.Value <> "ESSW (mecânica)" And FormularioDados.tipoestrutura.Value <> "ESSW (elétrica)" And FormularioDados.nrmodulos.Value <> "1 Módulo" And FormularioDados.nrmodulos.Value <> "2 Módulos" And FormularioDados.tipoestrutura.Value <> "Serviço Engenharia" Then
    MsgBox "Largura do Módulo 3 não informada. Favor verificar.", vbOKOnly
    Exit Sub
End If

If FormularioDados.largmodulo4.Value = 0 And FormularioDados.tipoestrutura.Value <> "Container Solar" And FormularioDados.tipoestrutura.Value <> "Skid (mecânica)" And FormularioDados.tipoestrutura.Value <> "Skid (com elétrica)" And FormularioDados.tipoestrutura.Value <> "Pilotis" And FormularioDados.tipoestrutura.Value <> "ESSW (mecânica)" And FormularioDados.tipoestrutura.Value <> "ESSW (elétrica)" And FormularioDados.nrmodulos.Value <> "1 Módulo" And FormularioDados.nrmodulos.Value <> "2 Módulos" And FormularioDados.nrmodulos.Value <> "3 Módulos" And FormularioDados.tipoestrutura.Value <> "Serviço Engenharia" Then
    MsgBox "Largura do Módulo 4 não informada. Favor verificar.", vbOKOnly
    Exit Sub
End If

If FormularioDados.largmodulo4.Value = 0 And FormularioDados.tipoestrutura.Value <> "Container Solar" And FormularioDados.tipoestrutura.Value <> "Skid (mecânica)" And FormularioDados.tipoestrutura.Value <> "Skid (com elétrica)" And FormularioDados.tipoestrutura.Value <> "Pilotis" And FormularioDados.tipoestrutura.Value <> "ESSW (mecânica)" And FormularioDados.tipoestrutura.Value <> "ESSW (elétrica)" And FormularioDados.nrmodulos.Value <> "1 Módulo" And FormularioDados.nrmodulos.Value <> "2 Módulos" And FormularioDados.nrmodulos.Value <> "3 Módulos" And FormularioDados.nrmodulos.Value <> "4 Módulos" And FormularioDados.tipoestrutura.Value <> "Serviço Engenharia" Then
    MsgBox "Largura do Módulo 5 não informada. Favor verificar.", vbOKOnly
    Exit Sub
End If

If FormularioDados.largmodulo4.Value = 0 And FormularioDados.tipoestrutura.Value <> "Container Solar" And FormularioDados.tipoestrutura.Value <> "Skid (mecânica)" And FormularioDados.tipoestrutura.Value <> "Skid (com elétrica)" And FormularioDados.tipoestrutura.Value <> "Pilotis" And FormularioDados.tipoestrutura.Value <> "ESSW (mecânica)" And FormularioDados.tipoestrutura.Value <> "ESSW (elétrica)" And FormularioDados.nrmodulos.Value <> "1 Módulo" And FormularioDados.nrmodulos.Value <> "2 Módulos" And FormularioDados.nrmodulos.Value <> "3 Módulos" And FormularioDados.nrmodulos.Value <> "4 Módulos" And FormularioDados.nrmodulos.Value <> "5 Módulos" And FormularioDados.tipoestrutura.Value <> "Serviço Engenharia" Then
    MsgBox "Largura do Módulo 6 não informada. Favor verificar.", vbOKOnly
    Exit Sub
End If

If FormularioDados.largmodulo4.Value = 0 And FormularioDados.tipoestrutura.Value <> "Container Solar" And FormularioDados.tipoestrutura.Value <> "Skid (mecânica)" And FormularioDados.tipoestrutura.Value <> "Skid (com elétrica)" And FormularioDados.tipoestrutura.Value <> "Pilotis" And FormularioDados.tipoestrutura.Value <> "ESSW (mecânica)" And FormularioDados.tipoestrutura.Value <> "ESSW (elétrica)" And FormularioDados.nrmodulos.Value <> "1 Módulo" And FormularioDados.nrmodulos.Value <> "2 Módulos" And FormularioDados.nrmodulos.Value <> "3 Módulos" And FormularioDados.nrmodulos.Value <> "4 Módulos" And FormularioDados.nrmodulos.Value <> "5 Módulos" And FormularioDados.nrmodulos.Value <> "6 Módulos" And FormularioDados.tipoestrutura.Value <> "Serviço Engenharia" Then
    MsgBox "Largura do Módulo 7 não informada. Favor verificar.", vbOKOnly
    Exit Sub
End If

If FormularioDados.largmodulo4.Value = 0 And FormularioDados.tipoestrutura.Value <> "Container Solar" And FormularioDados.tipoestrutura.Value <> "Skid (mecânica)" And FormularioDados.tipoestrutura.Value <> "Skid (com elétrica)" And FormularioDados.tipoestrutura.Value <> "Pilotis" And FormularioDados.tipoestrutura.Value <> "ESSW (mecânica)" And FormularioDados.tipoestrutura.Value <> "ESSW (elétrica)" And FormularioDados.nrmodulos.Value <> "1 Módulo" And FormularioDados.nrmodulos.Value <> "2 Módulos" And FormularioDados.nrmodulos.Value <> "3 Módulos" And FormularioDados.nrmodulos.Value <> "4 Módulos" And FormularioDados.nrmodulos.Value <> "5 Módulos" And FormularioDados.nrmodulos.Value <> "6 Módulos" And FormularioDados.nrmodulos.Value <> "7 Módulos" And FormularioDados.tipoestrutura.Value <> "Serviço Engenharia" Then
    MsgBox "Largura do Módulo 8 não informada. Favor verificar.", vbOKOnly
    Exit Sub
End If
If FormularioDados.tipomaq.Value = "" Then
    MsgBox "Tipo de máquina de ar condicionado não selecionado. Favor verificar.", vbOKOnly
    Exit Sub
End If

If FormularioDados.tipomaq.Value <> "Não possui" And FormularioDados.tipomaq.Value <> "Não aplicável" Then
    If FormularioDados.qtdmaq.Value = "" Or FormularioDados.qtdmaq.Value = 0 Then
        MsgBox "Quantidade de máquinas de ar condicionado não informado. Favor verificar.", vbOKOnly
        Exit Sub
    End If
End If

If FormularioDados.incendio.Value = "" Then
    MsgBox "Tipo de sistema de incêndio não selecionado. Favor verificar.", vbOKOnly
    Exit Sub
End If

If FormularioDados.seguranca.Value = "" Then
    MsgBox "Tipo de sistema de segurança não selecionado. Favor verificar.", vbOKOnly
    Exit Sub
End If

If FormularioDados.complexidade.Value = "" Then
    MsgBox "Complexidade dos equipamentos não selecionada. Favor verificar.", vbOKOnly
    Exit Sub
End If

If FormularioDados.nrcolunas.Value = "" Then
    MsgBox "Número de colunas não informado. Favor verificar.", vbOKOnly
    Exit Sub
End If

'RETIRADO PAINEL DE INTERLIGAÇÃO
'If FormularioDados.paineisint.Value = "" Then
'    MsgBox "Quantidade de painéis de interligação não informada. Favor verificar.", vbOKOnly
'    Exit Sub
'End If

If FormularioDados.tipomaq.Value = "Roof Top" And FormularioDados.Betim1310.Value = True Then
    MsgBox "Betim não está configurado para planejar projetos com ar condicionado Roof Top (casa de máquinas).", vbOKOnly
    Exit Sub
End If

If FormularioDados.tipomaq.Value = "Roof Top" And FormularioDados.nrmodulos.Value = "1 Módulo" Then
    MsgBox "Eletrocentros com tipo de Ar Condicionado Roof Top necessitam de casa máquinas (módulo adicional), porém foi informado que o eletrocentro possui somente um módulo. Favor verificar.", vbOKOnly
    Exit Sub
End If

If FormularioDados.txtPEP.Value = "" And (FormularioDados.chkCriarDRs.Value = True Or FormularioDados.chkMaterial.Value = True Or FormularioDados.chkPlanejar.Value = True Or FormularioDados.chkCamposUsuario.Value = True) Then
    MsgBox "PEP não informado. Favor verificar.", vbOKOnly
    Exit Sub
End If

If FormularioDados.txtDataInicio.Value = "" And FormularioDados.chkPlanejar.Value = True Then
    MsgBox "Data Início do Projeto não informada. Favor verificar.", vbOKOnly
    Exit Sub
End If

If FormularioDados.txtMaterial.Value = "" And FormularioDados.chkMaterial.Value = True And FormularioDados.tipoestrutura.Value <> "Serviço Engenharia" Then
    MsgBox "Material não informado. Favor verificar.", vbOKOnly
    Exit Sub
End If

If FormularioDados.txtCliente.Value = "" And FormularioDados.chkCamposUsuario.Value = True Then
    MsgBox "Cliente não informado. Favor verificar.", vbOKOnly
    Exit Sub
End If

If FormularioDados.cmbPlanejador.Value = "" And FormularioDados.chkCamposUsuario.Value = True Then
    MsgBox "Planejador não informado. Favor verificar.", vbOKOnly
    Exit Sub
End If

If FormularioDados.txtDataOV.Value = "" And FormularioDados.chkCamposUsuario.Value = True Then
    MsgBox "Data OV não informada. Favor verificar.", vbOKOnly
    Exit Sub
End If

If FormularioDados.txtValor.Value = "" And FormularioDados.chkCamposUsuario.Value = True And FormularioDados.tipoestrutura.Value <> "ESSW (elétrica)" And FormularioDados.tipoestrutura.Value <> "Serviço Engenharia" Then
    MsgBox "Valor do projeto (mecânica) não informado. Favor verificar.", vbOKOnly
    Exit Sub
End If

If FormularioDados.txtValorEletr.Value = "" And FormularioDados.chkCamposUsuario.Value = True And FormularioDados.tipoestrutura.Value <> "ESSW (mecânica)" And FormularioDados.tipoestrutura.Value <> "Pilotis" And FormularioDados.tipoestrutura.Value <> "Skid (mecânica)" Then
    MsgBox "Valor do projeto (elétrica) não informado. Favor verificar.", vbOKOnly
    Exit Sub
End If

If FormularioDados.txtPEP.Value <> "" And (Mid(FormularioDados.txtPEP.Value, 4, 1) <> "-" Or Mid(FormularioDados.txtPEP.Value, 12, 1) <> "-") Then
    MsgBox "PEP informado não possui formato válido. Favor verificar.", vbOKOnly
    Exit Sub
End If

If FormularioDados.txtDataOV.Value <> "" And (Mid(FormularioDados.txtDataOV.Value, 3, 1) <> "." Or Mid(FormularioDados.txtDataOV.Value, 6, 1) <> ".") Then
    MsgBox "Data da OV informada não possui formato válido (dd.mm.aaaa). Favor verificar.", vbOKOnly
    Exit Sub
End If

If FormularioDados.txtDataInicio.Value <> "" And (Mid(FormularioDados.txtDataInicio.Value, 3, 1) <> "." Or Mid(FormularioDados.txtDataInicio.Value, 6, 1) <> ".") Then
    MsgBox "Data de início informada não possui formato válido (dd.mm.aaaa). Favor verificar.", vbOKOnly
    Exit Sub
End If

If FormularioDados.nrmodulos.Value <> "1 Módulo" And FormularioDados.nrmodulos.Value <> "2 Módulos" And FormularioDados.nrmodulos.Value <> "3 Módulos" And FormularioDados.Betim1310.Value = True Then
    MsgBox "Qtd Módulos superior a 3 Módulos, não foi previsto para Betim. Favor verificar.", vbOKOnly
    Exit Sub
End If

FormularioDados.Hide
    
'Chamar programa para cálculo de tempos
Select Case FormularioDados.tipoestrutura.Value
    Case "Skid (mecânica)"
        Call tempos_skid
    Case "Pilotis"
        Call tempos_pilotis
    Case "Skid (com elétrica)"
        Call tempos_skid_eletrica
    Case "ESSW (elétrica)"
        Call tempos_ESSW_eletrica
    Case "Serviço Engenharia"
        Call tempos_servico_engenharia
    Case Else
        Call calcula
End Select

'ADICIONA AS ULTIMAS OPERAÇÕES COM TEMPOS FIXOS
Call operacoes_condicionais

'If FormularioDados.tipoestrutura.Value <> "Skid (mecânica)" And FormularioDados.tipoestrutura.Value <> "Skid (com elétrica)" And FormularioDados.tipoestrutura.Value <> "ESSW (elétrica)" And FormularioDados.tipoestrutura.Value <> "Pilotis" Then
'    Call calcula
'Else
'    If FormularioDados.tipoestrutura.Value = "Skid (mecânica)" Then
'        Call tempos_skid
'    Else
'        If FormularioDados.tipoestrutura.Value = "Pilotis" Then
'            Call tempos_pilotis
'        Else
'            If FormularioDados.tipoestrutura.Value = "Skid (com elétrica)" Then
'                Call tempos_skid_eletrica
'            Else
'                Call tempos_ESSW_eletrica
'            End If
'        End If
'    End If
'End If

'Chamar demais instruções
Call definir_template
erro_DR = False

If FormularioDados.chkCriarDRs.Value = True Then
    erro_DR = True
    Call criar_DRs
End If
If erro_DR = False Then
    If FormularioDados.chkPlanejar.Value = True Then Call planejar
    If FormularioDados.chkCamposUsuario.Value = True Then Call campos_usuario
    If FormularioDados.chkMaterial.Value = True Then Call amarrar_material
Else
    MsgBox "Erro na criação dos diagramas de rede. Execução interrompida.", vbOKOnly
End If

ThisWorkbook.Application.StatusBar = ""
ThisWorkbook.Application.StatusBar = False

End Sub
```

#### `tempos_servico_engenharia` (Sub)

```vb
Sub tempos_servico_engenharia()

Application.ScreenUpdating = False

If FormularioDados.txtPEP.Value <> "" Then
    ThisWorkbook.Application.StatusBar = "PEP " & FormularioDados.txtPEP.Value & ": Calculando tempos"
Else
    ThisWorkbook.Application.StatusBar = "Calculando tempos"
End If

ActiveSheet.Unprotect

'Limpa aba de resultados
Rows("4:245").Select
Selection.Delete Shift:=xlUp

'Insere template completo
template.Activate

ActiveSheet.Unprotect

Range("AQ1:AU21").Select
Selection.Copy

resultado.Activate
Range("A3").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False

template.Activate
ActiveSheet.Protect

'Formata tabela de tarefas
resultado.Activate
Range("A4:E23").Select

Selection.Borders(xlDiagonalDown).LineStyle = xlNone
Selection.Borders(xlDiagonalUp).LineStyle = xlNone
With Selection.Borders(xlEdgeLeft)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlEdgeRight)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlInsideVertical)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlInsideHorizontal)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With

Range("A4:E23").Select

With Selection
    .VerticalAlignment = xlBottom
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = False
End With

Range("A4:E23").Select

With Selection
    .VerticalAlignment = xlBottom
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = False
End With

Range("A4:A23").Select
With Selection
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlBottom
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = False
End With
Range("C4:E23").Select
With Selection
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlBottom
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = False
End With


Range("A4:A23").Select
Selection.NumberFormat = "0000"
Range("C4:C23").Select
Selection.NumberFormat = "0.0"
Range("E4:E23").Select
Selection.NumberFormat = "0.0"

Range("F3").Value = "SERVICO ENG"

'Inserir tempos
tempos.Activate
ActiveSheet.Unprotect

'Procura linha correspondente
For contador = 4 To 13
    If Range("A" & contador).Value = FormularioDados.nrmodulos.Value Then
              
        If FormularioDados.chaparemov.Value = True Then
            Range("G" & contador).Value = "Sim"
        Else
            Range("G" & contador).Value = "Não"
        End If
        
        If FormularioDados.trafooleo.Value = True Then
            Range("H" & contador).Value = "Sim"
        Else
            Range("H" & contador).Value = "Não"
        End If
        
        If FormularioDados.chkPeDireito = True Then
            Range("I" & contador).Value = "Sim"
        Else
            Range("I" & contador).Value = "Não"
        End If
        
        If FormularioDados.chkEscadaPadrao = True Then
            Range("J" & contador).Value = "Sim"
        Else
            Range("J" & contador).Value = "Não"
        End If
        
        If FormularioDados.chkEscadaEspecial = True Then
            Range("K" & contador).Value = "Sim"
        Else
            Range("K" & contador).Value = "Não"
        End If
        
        If FormularioDados.chkPoraoCabos = True Then
            Range("L" & contador).Value = "Sim"
        Else
            Range("L" & contador).Value = "Não"
        End If
        
        If FormularioDados.chkPilotis = True Then
            Range("M" & contador).Value = "Sim"
        Else
            Range("M" & contador).Value = "Não"
        End If
        
        If FormularioDados.chkRedeDutos = True Then
            Range("N" & contador).Value = "Sim"
        Else
            Range("N" & contador).Value = "Não"
        End If
        
        If FormularioDados.chkFundoFalso = True Then
            Range("O" & contador).Value = "Sim"
        Else
            Range("O" & contador).Value = "Não"
        End If
        
        If FormularioDados.chkDutosBWW = True Then
            Range("P" & contador).Value = "Sim"
        Else
            Range("P" & contador).Value = "Não"
        End If
        
        If FormularioDados.chkCalhasPluviais = True Then
            Range("Q" & contador).Value = "Sim"
        Else
            Range("Q" & contador).Value = "Não"
        End If
        
        If FormularioDados.chk_dutoGases = True Then
            Range("R" & contador).Value = "Sim"
        Else
            Range("R" & contador).Value = "Não"
        End If

        'Grava dados em variáveis Eng Mec
        HorLOM = Range("T" & contador).Value
        DurLOM = Range("U" & contador).Value
        HorLMM = Range("V" & contador).Value
        DurLMM = Range("W" & contador).Value
        HorPBS = Range("X" & contador).Value
        DurPBS = Range("Y" & contador).Value
        HorPPA = Range("Z" & contador).Value
        DurPPA = Range("AA" & contador).Value
        HorPCI = Range("AB" & contador).Value
        DurPCI = Range("AC" & contador).Value
        HorPCE = Range("AD" & contador).Value
        DurPCE = Range("AE" & contador).Value
        HorPAC = Range("AF" & contador).Value
        DurPAC = Range("AG" & contador).Value
        HorLCA = Range("AH" & contador).Value
        DurLCA = Range("AI" & contador).Value
        HorLAA = Range("AJ" & contador).Value
        DurLAA = Range("AK" & contador).Value
        HorLAM = Range("AL" & contador).Value
        DurLAM = Range("AM" & contador).Value
        HorLMA = Range("AN" & contador).Value
        DurLMA = Range("AO" & contador).Value
        HorPTR = Range("AP" & contador).Value
        DurPTR = Range("AQ" & contador).Value
    End If
Next

'Insere tempos na aba Resultados
resultado.Activate

Range("E4").Value = HorLOM
Range("C4").Value = DurLOM
Range("E8").Value = HorLMM
Range("C8").Value = DurLMM
Range("E9").Value = HorPBS + HorPPA
Range("C9").Value = DurPBS + DurPPA
Range("E10").Value = HorPCI
Range("C10").Value = DurPCI
Range("E11").Value = HorPCE
Range("C11").Value = DurPCE
Range("E12").Value = HorPAC
Range("C12").Value = DurPAC
Range("E14").Value = HorLCA
Range("C14").Value = DurLCA
Range("E15").Value = HorLAA
Range("C15").Value = DurLAA
Range("E16").Value = HorLAM
Range("C16").Value = DurLAM
Range("E17").Value = HorLMA
Range("C17").Value = DurLMA
Range("E19").Value = HorPTR
Range("C19").Value = DurPTR

'Cria tabela de totais de horas
Range("L3").Select
ActiveCell.FormulaR1C1 = "ENG"
Range("M3").Select
Selection.FormulaArray = _
    "=SUM((R4C5:R160C5)*(R4C5:R160C5<>0.1)*(R4C1:R160C1<=702)*(R4C1:R160C1<>531))"
Selection.NumberFormat = "#,##0.0"
Range("L4").Select
ActiveCell.FormulaR1C1 = "TOTAL"
Range("M4").Select
Selection.FormulaArray = "=SUM(R[-1]C:R[-1]C)"
Selection.NumberFormat = "#,##0.0"
Range("L4:M4").Select
Selection.Font.Bold = True
Range("L3:M4").Select
Selection.Borders(xlDiagonalDown).LineStyle = xlNone
Selection.Borders(xlDiagonalUp).LineStyle = xlNone
With Selection.Borders(xlEdgeLeft)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlEdgeRight)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlInsideVertical)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlInsideHorizontal)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With

'Seleciona resultado
Range("A4:E4").Select
Range(Selection, Selection.End(xlDown)).Select


ActiveSheet.Protect

End Sub
```

#### `operacoes_condicionais` (Sub)

```vb
Sub operacoes_condicionais()

Dim lastrow As Long, i As Long, j As Long

resultado.Unprotect


lastrow = resultado.Cells(Rows.Count, 1).End(xlUp).Row
For i = lastrow To 4 Step -1

    '###### ADICIONA TAREFA DE EMISSÃO DE RELATÓRIO SEMPRE QUE HOUVER PINTURA #######
    If (resultado.Cells(i, 1).Value = "753" Or resultado.Cells(i, 1).Value = "749" Or resultado.Cells(i, 1).Value = "750") And Left(resultado.Cells(i, 2), 3) = "PIN" Then
        j = i
        If FormularioDados.Betim1310.Value = True Then 'PINTURA EXTERNA BETIM
            resultado.Cells(j, 1).Value = "0749"
            resultado.Cells(j, 2).Value = "ESU - Envio para Subcontratação - PIN"
            resultado.Cells(j, 3).Value = "1,0"
            resultado.Cells(j, 4).Value = "DIA"
            resultado.Cells(j, 5).Value = "0,1"
        End If
        Do
            j = j + 1
            If CInt(resultado.Cells(j, 1).Value) > 756 Then
                resultado.Rows(j).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                resultado.Cells(j, 1).Value = "0756"
                resultado.Cells(j, 2).Value = "ERE - Emissão de Relatório"
                resultado.Cells(j, 3).Value = "1,0"
                resultado.Cells(j, 4).Value = "DIA"
                resultado.Cells(j, 5).Value = "0,1"
            End If
        Loop While CInt(resultado.Cells(j, 1).Value) < 756 And j < lastrow
    End If
    
Next i

resultado.Protect

End Sub
```
