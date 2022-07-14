VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Projetos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Private Sub Worksheet_SelectionChange(ByVal Target As Range)
' ' runs when a sheet is changed
'    'if Atualizando then goto
'
'    linha = Target.Row
'    coluna = Target.Column
'    valor = Target.Value
'    ID = Cells(linha, 1).Value
'    Dim cor As Integer
'    cor = 1
'
'    On Error GoTo fim
'
'    linha0 = AchaLinha("ID", 1)
'    col_prioridade = AchaColuna("Prioridade", linha0)
'
'    If col_prioridade <> coluna Then Exit Sub
'
'corresp:
'    corresp_linha = AchaLinha(valor, col_prioridade, cor)
'    If corresp_linha = linha Then
'        cor = cor + 1
'        GoTo corresp
'    ElseIf corresp_linha = "" Then
'        Exit Sub
'    End If
'
'    corresp_id = Cells(corresp_linha, 1).Value
'
'    FiltroPrioridade
'
'    linha_fim = Cells(linha0, col_prioridade).End(xlDown).Row
'
'    corresp_linha = AchaLinha(corresp_id, 1)
'
'    For i = corresp_linha To linha_fim
'        xid = Cells(i, 1).Value
'        If xid = ID Then GoTo pula
'        Cells(i, col_prioridade).Value = valor + 1
'pula:
'    Next i
'
'    FiltroPrioridade
'fim:
'
'End Sub
'
''procura a linha do outro item que tem a mesma prioridade
''pega o id do projeto e verifica com o que está sendo alterado
''se for igual procura a segunda correspondencia
''se não encontrar sai da macro
''da correspondencia em diante renumera os itens incrementando 1 até último item da lista
''reordena lista
'
'
'Public Function AchaLinha(ByVal texto As String, ByVal coluna As Integer, _
'                            Optional ocorrencia As Integer = 1) As Integer
''texto = texto procurado
''coluna = coluna onde está procurando o texto
''ocorrencia = padrão primeira ocorrencia
'
'        linha = 1
'        cel = Cells(linha, coluna).Value
'        x = 0
'
'A:
'        While cel <> texto And linha < 100000
'            linha = linha + 1
'            cel = Cells(linha, coluna).Value
'        Wend
'
'        x = x + 1
'        If x <> ocorrencia Then
'            linha = linha + 1
'            cel = Cells(linha, coluna).Value
'            GoTo A
'
'        ElseIf linha = 100000 Then
'            'MsgBox "Não consegui encontrar o que você está procurando!", vbExclamation, "Atenção"
'            AchaLinha = ""
'            GoTo fim
'        End If
'
'    AchaLinha = linha
'fim:
'End Function
'Public Function AchaColuna(ByVal texto As String, ByVal linha As Integer) As Integer
''texto = texto procurado
''linha = linha onde está procurando o texto
'    coluna = 1
'    cel = Cells(linha, coluna).Value
'    While cel <> texto And coluna < 15000
'        coluna = coluna + 1
'        cel = Cells(linha, coluna).Value
'    Wend
'
'    If coluna = 15000 Then
'        'MsgBox "Não consegui encontrar o que você está procurando!", vbExclamation, "Atenção"
'        GoTo fim
'    End If
'
'    AchaColuna = coluna
'fim:
'End Function
'
'Sub FiltroPrioridade()
'
'    ActiveWorkbook.Worksheets("Projetos").ListObjects("TabelaProjetos").Sort. _
'        SortFields.Clear
'    ActiveWorkbook.Worksheets("Projetos").ListObjects("TabelaProjetos").Sort. _
'        SortFields.Add Key:=Range("TabelaProjetos[[#All],[Prioridade]]"), SortOn:= _
'        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
'    With ActiveWorkbook.Worksheets("Projetos").ListObjects("TabelaProjetos").Sort
'        .Header = xlYes
'        .MatchCase = False
'        .Orientation = xlTopToBottom
'        .SortMethod = xlPinYin
'        .Apply
'    End With
'End Sub
'
