Attribute VB_Name = "PROCESSOS"
Sub ADC_PROCESSO()
' Adiciona planilha de processos comum.

    
    Set ThisWB = ThisWorkbook
    Set ThisWS = ThisWB.ActiveSheet
    Set TemplateWS = ThisWB.Worksheets("PROCESSO")
    
    WSCount = ThisWB.Worksheets.Count
    j = 0
    
    For i = 1 To WSCount
        
        If ThisWB.Sheets(i).Visible = True Then: j = j + 1

    Next i
    
    VisWSCount = j
    j = 0
    
    For i = 1 To WSCount
        
        If ThisWB.Sheets(i).Visible = True Then: j = j + 1
        
        If j = VisWSCount Then
            
            LastWS = i
            Exit For
        End If
        
    Next i
    
    TemplateWS.Visible = True
    
    TemplateWS.Copy After:=Sheets(LastWS)
    
    Set AddedWS = ThisWB.ActiveSheet
    
    BLOQUEIO AddedWS
    
    TemplateWS.Visible = False
    
    AddedWS.Cells(14, 1).Select

End Sub

Sub CADASTRA_PROCESSO()

 
    sNovoProc = UCase(Range("E5").Value)
    sTipo = Range("E6").Value
    
    If sNovoProc = "" Or sTipo = "" Then
        MsgBox "Os campos 'Nome do Processo' e 'Método de controle' devem estar preenchidos", vbExclamation, "Atenção"
        Exit Sub
    End If
    
    Sheets("DADOS").Visible = True
    Sheets("DADOS").Select
    
    iLinha = 2
    sProcesso = Cells(iLinha, 2)
    
    While sProcesso <> ""
        If sProcesso = sNovoProc Then
            Sheets("DADOS").Visible = False
            Sheets("CADASTRO").Select
            MsgBox "Esse processo já está cadastrado", vbInformation, "Atenção"
            Exit Sub
        End If
     
     iLinha = iLinha + 1
     sProcesso = Cells(iLinha, 2).Value
    Wend
    
    Cells(iLinha, 2) = sNovoProc
    Cells(iLinha, 3) = sTipo
    
    ActiveWorkbook.Worksheets("DADOS").ListObjects("Tabela1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("DADOS").ListObjects("Tabela1").Sort.SortFields.Add2 _
        Key:=Range("Tabela1[[#All],[PROCESSOS]]"), SortOn:=xlSortOnValues, Order _
        :=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("DADOS").ListObjects("Tabela1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Sheets("DADOS").Visible = False
    Sheets("CADASTRO").Select
    MsgBox "Processo " & sNovoProc & " foi cadastrado com sucesso!", vbInformation, "Concluído!"
    
    Range("E5").Value = ""
    Range("E6").Value = ""
End Sub
