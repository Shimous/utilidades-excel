VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_ETM002 
   Caption         =   "Cadastro POS. ETM 002"
   ClientHeight    =   10395
   ClientLeft      =   45
   ClientTop       =   465
   ClientWidth     =   9270
   OleObjectBlob   =   "UF_ETM002.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_ETM002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cadastra_Click()
'Cadastro de ETM no banco de dados
    Dim sPos As String       ' Posição da ETM
    Dim sTS As String        ' Tratamento superficial
    Dim sMetoMed As String   ' Método de medição TS
    Dim sEspCI As String     ' LIE espessura da camada
    Dim sEspCS As String     ' LSE espessura da camada
    Dim sCV As String        ' Corrosão vermelha
    Dim sCB As String        ' Corrosão branca
    Dim sUn As String        ' Unidade de medida salt spray
    Dim sDTTi As String      ' LIE DTT
    Dim sDTTs As String      ' LSE DTT
    Dim sDTTiC As String      ' LIE DTT na cabeça e rosca
    Dim sDTTsC As String      ' LSE DTT na cabeça e rosca
    exportar = 0
'Rotinas de gravação dos valores nas variáveis
    
    sPos = PosicaoETM
    If sPos = "Pedro" Then
        Exit Sub
    End If

VerificaRepetido:
    'Verifica se a posição já está cadastrada
    Sheets("DADOS").Visible = True
    Sheets("DADOS").Select
    Linha = 2
    PosicaoCadastrada = Cells(Linha, 5).Value
    
    While PosicaoCadastrada <> ""
        If PosicaoCadastrada = sPos And exportar = 0 Then
            MsgBox sPos & " já está cadastrada", vbExclamation, "Atenção"
            Sheets("DADOS").Visible = False
            Sheets("CADASTRO").Select
            Exit Sub
        End If
        If PosicaoCadastrada = sPos And exportar = 77 Then
            MsgBox sPos & " já está cadastrada no Modelo 1", vbExclamation, "Atenção"
            Sheets("DADOS").Visible = False
            Sheets("CADASTRO").Select
            VerRepTrue = 1
            GoTo Reset
        End If
        If PosicaoCadastrada = sPos And exportar = 78 Then
            MsgBox sPos & " já está cadastrada no Modelo 1", vbExclamation, "Atenção"
            Sheets("DADOS").Visible = False
            Sheets("CADASTRO").Select
            GoTo Reset
        End If
        If PosicaoCadastrada = sPos And exportar = 8 Then
            MsgBox sPos & " já está cadastrada no Modelo 1.1", vbExclamation, "Atenção"
            Sheets("DADOS").Visible = False
            Sheets("CADASTRO").Select
            GoTo Reset
        End If
        Linha = Linha + 1
        PosicaoCadastrada = Cells(Linha, 5).Value
    Wend
    Sheets("DADOS").Visible = False
    If exportar <> 0 Then
        GoTo EscreveBD
    End If
    
    If UF_ETM002.TS.Value = "" Then
        MsgBox "Digite o nome do Tratamento superficial.", vbExclamation, "Atenção"
        Exit Sub
    End If
    
    sTS = UCase(UF_ETM002.TS.Value)
    
    sMetoMed = MetoMedTS
    If sMetoMed = "Pedro" Then
        Exit Sub
    End If
    
    sEspCI = EspessuraCamadaI
    If sEspCI = "Pedro" Then
        Exit Sub
    End If
    
    sEspCS = EspessuraCamadaS
    If sEspCS = "Pedro" Then
        Exit Sub
    End If
    
    sCV = CorrVer
    If sCV = "Pedro" Then
        Exit Sub
    End If
    
    sCB = CorrBra
    If sCB = "Pedro" Then
        Exit Sub
    End If
    
    sUn = UnidadeSP
    
    sDTTi = DTTinf
    If sDTTi = "Pedro" Then
        Exit Sub
    End If
    
    sObsExtras = obsSP.Value
    
    sDTTs = DTTsup
    If sDTTs = "Pedro" Then
        Exit Sub
    End If
    
    sDTTiC = DTTinfCabRos
    If sDTTiC = "Pedro" Then
        Exit Sub
    End If
    
    sDTTsC = DTTsupCabRos
    If sDTTsC = "Pedro" Then
        Exit Sub
    End If
    
EscreveBD:
'Rotina pra cadastrar dados no banco de dados
    Sheets("ETM").Visible = True
    Sheets("ETM").Select
    Dim Inicio As Range
    Set Inicio = Range("D1").End(xlDown).Offset(1, 0)
    iLinha1 = 0
    
    'Primeira linha TS
    Inicio.Value = sPos
    Inicio.Offset(iLinha1, 1).Value = "TRATAMENTO SUPERFICIAL"
    Inicio.Offset(iLinha1, 2).Value = sTS
    Inicio.Offset(iLinha1, 5).Value = sMetoMed
    Inicio.Offset(iLinha1, 6).Value = "2"
    Inicio.Offset(iLinha1, 7).Value = "LOTE"
    Inicio.Offset(iLinha1, 8).Value = "N/A"
    Inicio.Offset(iLinha1, 9).Value = "R.I.R."
    
    'Segunda linha Espessura da Camada
    If UF_ETM002.naEsp.Value = False Then
        iLinha1 = iLinha1 + 1
        Inicio.Offset(iLinha1, 0).Value = sPos
        Inicio.Offset(iLinha1, 1).Value = "ESPESSURA DA CAMADA"
        Inicio.Offset(iLinha1, 3).Value = sEspCI
        Inicio.Offset(iLinha1, 4).Value = sEspCS
        Inicio.Offset(iLinha1, 5).Value = "MEDIDOR DE CAMADA"
        Inicio.Offset(iLinha1, 6).Value = "2"
        Inicio.Offset(iLinha1, 7).Value = "LOTE"
        Inicio.Offset(1, 8).Value = "µm"
        Inicio.Offset(1, 9).Value = "R.I.R."
    End If
    
    'Terceira linha Aspecto Visual
    iLinha1 = iLinha1 + 1
    Inicio.Offset(iLinha1, 0).Value = sPos
    Inicio.Offset(iLinha1, 1).Value = "ASPECTO VISUAL"
    Inicio.Offset(iLinha1, 2).Value = "ISENTO DE FALHAS, MANCHAS E OXIDAÇÕES"
    Inicio.Offset(iLinha1, 5).Value = "VISUAL"
    Inicio.Offset(iLinha1, 6).Value = "2"
    Inicio.Offset(iLinha1, 7).Value = "LOTE"
    Inicio.Offset(iLinha1, 8).Value = "N/A"
    Inicio.Offset(iLinha1, 9).Value = "R.I.R."

    'Quarta linha Salt Spray
    If UF_ETM002.naSP.Value = False Then
        iLinha1 = iLinha1 + 1
        Inicio.Offset(iLinha1, 0).Value = sPos
        Inicio.Offset(iLinha1, 1).Value = "SALT SPRAY"
        
        If UF_ETM002.ToggleButton4.Value = True Then
            If sUn = "SEMANAS" Then
                Inicio.Offset(iLinha1, 2).Value = sCB & " " & sUn & " ISENTO DE CORROSÃO BRANCA" & sObsExtras
            Else
                Inicio.Offset(iLinha1, 2).Value = sCB & sUn & " ISENTO DE CORROSÃO BRANCA" & sObsExtras
            End If
        End If
        
        If UF_ETM002.ToggleButton5.Value = True Then
            If sUn = "SEMANAS" Then
                Inicio.Offset(iLinha1, 2).Value = sCV & " " & sUn & " ISENTO DE CORROSÃO VERMELHA" & sObsExtras
            Else
                Inicio.Offset(iLinha1, 2).Value = sCV & sUn & " ISENTO DE CORROSÃO VERMELHA" & sObsExtras
            End If
        End If
        If sUn = "SEMANAS" Then
            Inicio.Offset(iLinha1, 2).Value = sCB & " " & sUn & " ISENTO DE CORROSÃO BRANCA, " & sCV & " " & sUn & " ISENTO DE CORROSÃO VERMELHA" & sObsExtras
        Else
            Inicio.Offset(iLinha1, 2).Value = sCB & sUn & " ISENTO DE CORROSÃO BRANCA, " & sCV & sUn & " ISENTO DE CORROSÃO VERMELHA" & sObsExtras
        End If
        Inicio.Offset(iLinha1, 5).Value = "MÁQ. DE SALT SPRAY"
        Inicio.Offset(iLinha1, 6).Value = "5"
        Inicio.Offset(iLinha1, 7).Value = "LOTE"
        Inicio.Offset(iLinha1, 8).Value = sUn
        Inicio.Offset(iLinha1, 9).Value = "R.I.R."
    End If
    
    'Quinta linha DTT
    If UF_ETM002.naDTT.Value = False Then
        iLinha1 = iLinha1 + 1
        Inicio.Offset(iLinha1, 0).Value = sPos
        If UF_ETM002.ToggleButton9.Value = True Then
            Inicio.Offset(iLinha1, 1).Value = "ENSAIO DTT GERAL"
        Else
            Inicio.Offset(iLinha1, 1).Value = "ENSAIO DTT"
        End If
        Inicio.Offset(iLinha1, 3).Value = sDTTi
        Inicio.Offset(iLinha1, 4).Value = sDTTs
        Inicio.Offset(iLinha1, 5).Value = "MÁQ. DE DTT"
        Inicio.Offset(iLinha1, 6).Value = "5"
        Inicio.Offset(iLinha1, 7).Value = "LOTE"
        Inicio.Offset(iLinha1, 8).Value = "µGes"
        Inicio.Offset(iLinha1, 9).Value = "R.I.R."
    End If
    
    'Sexta linha DTT cabeça e rosca
    If UF_ETM002.naDTT.Value = False And UF_ETM002.ToggleButton9.Value = True Then
        iLinha1 = iLinha1 + 1
        Inicio.Offset(iLinha1, 0).Value = sPos
        Inicio.Offset(iLinha1, 1).Value = "ENSAIO DTT ROSCA E CABEÇA"
        Inicio.Offset(iLinha1, 3).Value = sDTTiC
        Inicio.Offset(iLinha1, 4).Value = sDTTsC
        Inicio.Offset(iLinha1, 5).Value = "MÁQ. DE DTT"
        Inicio.Offset(iLinha1, 6).Value = "5"
        Inicio.Offset(iLinha1, 7).Value = "LOTE"
        Inicio.Offset(iLinha1, 8).Value = "µG"
        Inicio.Offset(iLinha1, 9).Value = "R.I.R."
    End If
    
    Sheets("DADOS").Visible = True
    Sheets("DADOS").Select
    Set DadosETM = Range("E1").End(xlDown).Offset(1, 0)
    DadosETM.Value = sPos
    
    ActiveWorkbook.Worksheets("DADOS").ListObjects("Tabela4").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("DADOS").ListObjects("Tabela4").Sort.SortFields.Add _
        Key:=Range("Tabela4[[#All],[ETM 002]]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("DADOS").ListObjects("Tabela4").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Sheets("DADOS").Visible = False
    
    Sheets("ETM").Visible = False
    Sheets("CADASTRO").Select
    
  
    If exportar = 0 Then
        MsgBox sPos & " cadastrada com sucesso!", vbInformation, "Sucesso"
        If ThisWorkbook.Name = "1. MODELO DE PLANO DE CONTROLE.xlsm" Then
            
            exportar = MsgBox("Deseja gravar esta posição no Modelo 1.1 (D) de plano de controle?", vbQuestion + vbYesNo, "Exportar")
            If exportar = 6 Then
                exportar = 8
                Workbooks.Open ("\\192.168.1.19\Dados\Arquivos\ENGENHARIA\13. Planos de Controle\1.1 MODELO DE PLANO DE CONTROLE (D).xlsm")
                GoTo VerificaRepetido
            Else
                GoTo Reset
            End If
        End If
        
         If ThisWorkbook.Name = "1.1 MODELO DE PLANO DE CONTROLE (D).xlsm" Then
            exportar = MsgBox("Deseja gravar esta posição no Modelo 1. de plano de controle?", vbQuestion + vbYesNo, "Exportar")
            If exportar = 6 Then
                Workbooks.Open ("\\192.168.1.19\Dados\Arquivos\ENGENHARIA\13. Planos de Controle\1. MODELO DE PLANO DE CONTROLE.xlsm")
                exportar = 78
                GoTo VerificaRepetido
            Else
                GoTo Reset
            End If
        End If
        
        exportar = MsgBox("Deseja gravar esta posição nos Modelos de plano de controle?", vbQuestion + vbYesNo, "Exportar")
        If exportar = 6 Then
            Workbooks.Open ("\\192.168.1.19\Dados\Arquivos\ENGENHARIA\13. Planos de Controle\1. MODELO DE PLANO DE CONTROLE.xlsm")
            exportar = 77
            GoTo VerificaRepetido
        End If
        If exportar = 7 Then
            GoTo Reset
        End If
    End If

    
    If exportar = 78 Then
        Sheets("INSTRUÇÕES").Select
        ActiveWorkbook.Save
        ActiveWorkbook.Close
        MsgBox sPos & " cadastrada com sucesso no Modelo 1!", vbInformation, "Sucesso"
        GoTo Reset
        
    End If
    
ModeloD:
    If exportar = 77 Then
        Sheets("INSTRUÇÕES").Select
        ActiveWorkbook.Save
        ActiveWorkbook.Close
        If VerRepTrue = 1 Then GoTo CadastraD
        MsgBox sPos & " cadastrada com sucesso no Modelo 1!", vbInformation, "Sucesso"
CadastraD:
        Workbooks.Open ("\\192.168.1.19\Dados\Arquivos\ENGENHARIA\13. Planos de Controle\1.1 MODELO DE PLANO DE CONTROLE (D).xlsm")
        exportar = 8
        GoTo VerificaRepetido
        
    End If

    If exportar = 8 Then
        Sheets("INSTRUÇÕES").Select
        ActiveWorkbook.Save
        ActiveWorkbook.Close
        MsgBox sPos & " cadastrada com sucesso no Modelo 1.1!", vbInformation, "Sucesso"
        exportar = 0
    End If
    
Reset:
    With UF_ETM002
        .Pos.Value = ""
        .TS.Value = ""
        .OutroMetoMed.Value = ""
        .TextBox3.Value = ""
        .TextBox4.Value = ""
        .TextBox5.Value = ""
        .TextBox6.Value = ""
        .TextBox7.Value = ""
        .TextBox8.Value = ""
        .TextBox9.Value = ""
        .TextBox10.Value = ""
        .Hide
    End With
End Sub
Function DTTsupCabRos()
    If UF_ETM002.naDTT.Value = True Or UF_ETM002.naLSEdttC.Value = True Then
        DTTsupCabRos = ""
        Exit Function
    End If
    If UF_ETM002.naLSEdttC.Value = False And UF_ETM002.TextBox9.Value = "" Then
        MsgBox "Especifique LSE de coeficiente de atrito na cabeça e rosca, ou selecione o botão N/A ao lado do campo LSE", vbExclamation, "Atenção"
        DTTsupCabRos = "Pedro"
        Exit Function
    End If
    
    DTTsupCabRos = UF_ETM002.TextBox9.Value
End Function

Function DTTinfCabRos()
    If UF_ETM002.naDTT.Value = True Or UF_ETM002.naLIEdttC.Value = True Then
        DTTinfCabRos = ""
        Exit Function
    End If
    
    If UF_ETM002.naLIEdttC.Value = False And UF_ETM002.TextBox10.Value = "" Then
        MsgBox "Especifique LIE de coeficiente de atrito na cabeça e rosca, ou selecione o botão N/A ao lado do campo LIE", vbExclamation, "Atenção"
        DTTinfCabRos = "Pedro"
        Exit Function
    End If
    
    DTTinfCabRos = UF_ETM002.TextBox10.Value
End Function


Function DTTsup() As String
    If UF_ETM002.naDTT.Value = True Or UF_ETM002.naLSEdtt.Value = True Then
        DTTsup = ""
        Exit Function
    End If
    
    If UF_ETM002.naLSEdtt.Value = False And UF_ETM002.TextBox6.Value = "" Then
        MsgBox "Especifique LSE de coeficiente de atrito, ou selecione o botão N/A ao lado do campo LSE", vbExclamation, "Atenção"
        DTTsup = "Pedro"
        Exit Function
    End If
    DTTsup = UF_ETM002.TextBox6.Value
    
End Function
Function DTTinf() As String
    If UF_ETM002.naDTT.Value = True Or UF_ETM002.naLIEdtt.Value = True Then
        DTTinf = ""
        Exit Function
    End If
    
    If UF_ETM002.naLIEdtt.Value = False And UF_ETM002.TextBox5.Value = "" Then
        MsgBox "Especifique LIE de coeficiente de atrito, ou selecione o botão N/A ao lado do campo LIE", vbExclamation, "Atenção"
        DTTinf = "Pedro"
        Exit Function
    End If
    DTTinf = UF_ETM002.TextBox5.Value
    
End Function

Function UnidadeSP() As String
    If UF_ETM002.horas.Value = True Then
        UnidadeSP = "H"
    Else
        UnidadeSP = "SEMANAS"
    End If

End Function

Function CorrBra() As String
    If UF_ETM002.naSP.Value = True Then
        CorrBra = ""
        Exit Function
    End If
    
    If UF_ETM002.ToggleButton5.Value = True Then
        CorrBra = ""
        Exit Function
    End If
    
    If UF_ETM002.ToggleButton5.Value = False And UF_ETM002.TextBox8.Value = "" Then
        MsgBox "Especifique um valor para Corrosão branca, ou selecione o botão N/A ao lado do campo CB", vbExclamation, "Atenção"
        CorrBra = "Pedro"
        Exit Function
    End If
    
    If UF_ETM002.ToggleButton5.Value = False Then
        CorrBra = UF_ETM002.TextBox8.Value
    End If
    
End Function

Function CorrVer() As String
    If UF_ETM002.naSP.Value = True Then
        CorrVer = ""
        Exit Function
    End If
    
    If UF_ETM002.ToggleButton4.Value = True Then
        CorrVer = ""
        Exit Function
    End If
    
    If UF_ETM002.ToggleButton4.Value = False And UF_ETM002.TextBox7.Value = "" Then
        MsgBox "Especifique um valor para Corrosão vermelha, ou selecione o botão N/A ao lado do campo CV", vbExclamation, "Atenção"
        CorrVer = "Pedro"
        Exit Function
    End If
    
    If UF_ETM002.ToggleButton4.Value = False Then
        CorrVer = UF_ETM002.TextBox7.Value
    End If
End Function
Function EspessuraCamadaS() As String
'Grava o LSE de espessura de camada
    If UF_ETM002.naEsp.Value = True Or UF_ETM002.naLSEec.Value = True Then
        EspessuraCamadaS = ""
        Exit Function
    End If
    
    If UF_ETM002.TextBox4.Value = "" Then
            MsgBox "Especifique o LSE de espessura de camada, ou selecione o botão N/A ao lado do campo LSE", vbExclamation, "Atenção"
        EspessuraCamadaS = "Pedro"
        Exit Function
    End If
    
    EspessuraCamadaS = UF_ETM002.TextBox4.Value
    
    If UF_ETM002.naLIEec.Value = True And UF_ETM002.naLSEec.Value = True Then
        resposta = MsgBox("Apenas o limite inferior de espessura de camada foi informado, deseja prosseguir?", vbQuestion + vbYesNo, "Atenção")
        If resposta = 6 Then
            Exit Function
        Else
            EspessuraCamadaS = "Pedro"
            Exit Function
        End If
    End If
    
    If UF_ETM002.naLIEec.Value = True And UF_ETM002.naLSEec.Value = True Then
    
        resposta = MsgBox("Apenas o limite superior de espessura de camada foi informado, deseja prosseguir?", vbQuestion + vbYesNo, "Atenção")
        If resposta = 6 Then
            Exit Function
        Else
            EspessuraCamadaS = "Pedro"
            Exit Function
        End If
    End If
    


    
End Function

Function EspessuraCamadaI() As String
'Grava o LIE de espessura de camada
    If UF_ETM002.naEsp.Value = True Or UF_ETM002.naLIEec.Value = True Then
        EspessuraCamadaI = ""
        Exit Function
    End If
    
    If UF_ETM002.TextBox3.Value = "" Then
        MsgBox "Especifique o LIE de espessura de camada, ou selecione o botão N/A ao lado do campo LIE", vbExclamation, "Atenção"
        EspessuraCamadaI = "Pedro"
        Exit Function
    End If
    
    EspessuraCamadaI = UF_ETM002.TextBox3.Value
End Function

Function PosicaoETM() As String
'Grava a posição da ETM com padrão de 3 dígitos
    If Pos.Value = "" Then
        MsgBox "Digite o número da posição de ETM desejada.", vbExclamation, "Atenção"
        PosicaoETM = "Pedro"
        Exit Function
    End If
    
    If Len(Pos.Value) = 2 Then
        PosicaoETM = "ETM 002 POS.0" & Pos.Value
        Exit Function
    End If
    
    If Len(Pos.Value) = 1 Then
        PosicaoETM = "ETM 002 POS.00" & Pos.Value
        Exit Function
    End If
    
    PosicaoETM = "ETM 002 POS." & Pos.Value
    
End Function

Function MetoMedTS() As String
    If CONF_CERT.Value = True Then
        MetoMedTS = "CONF. CERT. FORNECEDOR"
        Exit Function
    End If

    If VISUAL.Value = True Then
        MetoMedTS = "VISUAL"
        Exit Function
    End If

    If OUTRO.Value = True Then
        If OutroMetoMed.Value = "" Then
            MsgBox "Digite o método de avaliação do TS, ou selecione outra Tecnica de avaliação.", vbExclamation, "Atenção"
            sMetoMed = "Pedro"
            Exit Function
        End If
        MetoMedTS = UCase(OutroMetoMed.Value)
        Exit Function
    End If
    
End Function


'Scripts de objetos ==========================================================================================================================
Private Sub CONF_CERT_Click()
    OutroMetoMed.Visible = OUTRO.Value
End Sub


Private Sub naDTT_Click()
'Habiliba ou desabilita quadro de DTT
    If naDTT.Value = True Then
        UF_ETM002.Frame1.Visible = False
        UF_ETM002.DTT.Visible = True
        UF_ETM002.naDTT.Font.Size = 8
        UF_ETM002.naDTT.Caption = "Habilitar"
    Else
        UF_ETM002.Frame1.Visible = True
        UF_ETM002.DTT.Visible = False
        UF_ETM002.naDTT.Font.Size = 12
        UF_ETM002.naDTT.Caption = "N/A"
    End If
End Sub

Private Sub naSP_Click()
'Habiliba ou desabilita quadro de Salt Spray
    If naSP.Value = True Then
        UF_ETM002.Frame2.Visible = False
        UF_ETM002.Salt.Visible = True
        UF_ETM002.naSP.Font.Size = 8
        UF_ETM002.naSP.Caption = "Habilitar"
    Else
        UF_ETM002.Frame2.Visible = True
        UF_ETM002.Salt.Visible = False
        UF_ETM002.naSP.Font.Size = 12
        UF_ETM002.naSP.Caption = "N/A"
    End If
End Sub

Private Sub naEsp_Click()
'Habiliba ou desabilita quadro de Espessura da camada
    If naEsp.Value = True Then
        UF_ETM002.Frame3.Visible = False
        UF_ETM002.EspCam.Visible = True
        UF_ETM002.naEsp.Font.Size = 8
        UF_ETM002.naEsp.Caption = "Habilitar"
    Else
        UF_ETM002.Frame3.Visible = True
        UF_ETM002.EspCam.Visible = False
        UF_ETM002.naEsp.Font.Size = 12
        UF_ETM002.naEsp.Caption = "N/A"
    End If
End Sub


Private Sub OUTRO_Click()
    OutroMetoMed.Visible = OUTRO.Value
End Sub


Private Sub ToggleButton4_Click()
'Habiliba ou desabilita CV
    If ToggleButton4.Value = True Then
        UF_ETM002.TextBox7.Enabled = False
        UF_ETM002.Label7.Enabled = False
    Else
        UF_ETM002.TextBox7.Enabled = True
        UF_ETM002.Label7.Enabled = True
    End If

End Sub

Private Sub ToggleButton5_Click()
'Habiliba ou desabilita CB
    If ToggleButton5.Value = True Then
        UF_ETM002.TextBox8.Enabled = False
        UF_ETM002.Label8.Enabled = False
    Else
        UF_ETM002.TextBox8.Enabled = True
        UF_ETM002.Label8.Enabled = True
    End If
    
End Sub

Private Sub ToggleButton6_Click()
    obsSP.Visible = ToggleButton6.Value
    
End Sub

Private Sub ToggleButton9_Click()

    With UF_ETM002
        .TextBox10.Visible = UF_ETM002.ToggleButton9.Value
        .TextBox9.Visible = UF_ETM002.ToggleButton9.Value
        .naLIEdttC.Visible = UF_ETM002.ToggleButton9.Value
        .naLSEdttC.Visible = UF_ETM002.ToggleButton9.Value
        .Label16.Visible = UF_ETM002.ToggleButton9.Value
        .Label17.Visible = UF_ETM002.ToggleButton9.Value
    End With
End Sub

Private Sub UserForm_Initialize()
    ToggleButton5.Value = True
    OutroMetoMed.Visible = OUTRO.Value
End Sub

Private Sub VISUAL_Click()
    OutroMetoMed.Visible = OUTRO.Value
End Sub
