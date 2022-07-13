VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EMAIL_CONFIG 
   Caption         =   "Solicitação de aprovação"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10320
   OleObjectBlob   =   "EMAIL_CONFIG.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "EMAIL_CONFIG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub UserForm_Activate()

    gerente1.Caption = Cells(AchaLinha("Aprovador", AchaColuna("Aprovador", 1)) + 1, 2).Value
    gerente1.Value = True
    gerente2.Caption = Cells(AchaLinha("Aprovador", AchaColuna("Aprovador", 1)) + 2, 2).Value
    
    
    linha0 = AchaLinha("ID", 1)
    linha = Cells(linha0, 1).End(xlDown).Row
    id = Cells(linha, 1).Value
    titulo = Cells(linha, 15).Value
    Solic = id & " - " & titulo
    
    While id <> "ID"
        Lista.AddItem (Solic)
        linha = linha - 1
        id = Cells(linha, 1).Value
        titulo = Cells(linha, 15).Value
        Solic = id & " - " & titulo
    Wend
    
End Sub
Private Sub Avancar_Click()
    Solic = Lista.Value
    If IsNull(Lista.Value) Then
        ok = MsgBox("Tem que selecionar uma solicitação né LERDÃO!", vbOKOnly + vbExclamation, "Atenção")
        GoTo fim
    End If
    Dim id As Integer
    
    linha0 = AchaLinha("ID", 1)
    id = Left(Solic, InStr(1, Solic, " ") - 1) + linha0
    nome = Cells(id, AchaColuna("Titulo", linha0)).Value
    titulo = "Solicitação de compras - " & nome
    
    If Cells(id, 8).Value = "" And Cells(id, 9).Value = "" Then
        valor = ""
    
    ElseIf Cells(id, 8).Value <> "" Then
        preco = Format(Cells(id, AchaColuna("Valor total (R$)", linha0)).Value, "#,###.00")
        valor = "<b>Valor:</b> R$" & preco
    
    Else
        preco = Format(Cells(id, AchaColuna("Valor total (U$)", linha0)).Value, "#,###.00")
        valor = "<b>Valor:</b> U$" & preco
        
    End If
         
    If gerente1.Value = True Then
        gerente = gerente1.Caption
    Else
        gerente = gerente2.Caption
    End If
    destinatario = Cells(AchaLinha(gerente, AchaColuna("Aprovador", 1)), AchaColuna("Email", 1)).Value
    
    If valor <> "" Then
        
    mensagem = _
        "<font size='11pt' face='Calibri'>" & _
        gerente & ", <br><br>" & _
        "Segue Nº do esboço referente à solicitação de compra <b><font color=#0066cc>" & nome & "</font></b>:<br><br>" & _
        "<b>Nº da chave: </b>" & Cells(id, 5).Value & "<br>" & _
        "<b>Nº do esboço: </b>" & Cells(id, 4).Value & "<br>" & _
        valor & "<br><br>" & _
        "Itens: " & RangeToHTML(Range(Cells(id, 3), Cells(id, 3))) & "<br><br>" & _
        "Aguardando aprovação.<br><br>" & _
        "Grato,</font>"
    
    Else
    
    mensagem = _
        "<font size='11pt' face='Calibri'>" & _
        gerente & ", <br><br>" & _
        "Segue Nº do esboço referente à solicitação de compra <b><font color=#0066cc>" & nome & "</font></b>:<br><br>" & _
        "<b>Nº da chave: </b>" & Cells(id, 5).Value & "<br>" & _
        "<b>Nº do esboço: </b>" & Cells(id, 4).Value & "<br><br>" & _
        "<b>Itens:</b> " & RangeToHTML(Range(Cells(id, 3), Cells(id, 3))) & "<br><br>" & _
        "Aguardando aprovação.<br><br>" & _
        "Grato,</font>"
    
    End If
    
    Set OutlookApp = CreateObject("Outlook.Application")
    Set OutlookMail = OutlookApp.CreateItem(0)
    
    OutlookMail.display
   
    With OutlookMail
      
      .To = destinatario
      .CC = "joao.gross@metaltork.com.br"
      .Subject = titulo
      .htmlBody = mensagem & OutlookMail.htmlBody
     End With
    
    If Cells(id, AchaColuna("Status", linha0)).Value = "Enviar" Or Cells(id, 6).Value = "" Then
        Cells(id, 6).Value = "Pendente"
    End If
    
fim:
        
End Sub

Private Sub UserForm_Terminate()
    ActiveSheet.Shapes.Range(Array("EMAIL")).Select
    Selection.ShapeRange.Shadow.Visible = msoFalse
    Range("A9").Select
End Sub

Function RangeToHTML(rng As Range)
' Changed by Ron de Bruin 28-Oct-2006
' Working in Office 2000-2016
    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook

    TempFile = Environ$("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

    'Copy the range and create a new workbook to past the data in
    rng.Copy
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
    End With

    'Publish the sheet to a htm file
    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         Filename:=TempFile, _
         Sheet:=TempWB.Sheets(1).Name, _
         Source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With

    'Read all data from the htm file into RangetoHTML
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangeToHTML = ts.readall
    ts.Close
    RangeToHTML = Replace(RangeToHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")

    'Close TempWB
    TempWB.Close savechanges:=False

    'Delete the htm file we used in this function
    Kill TempFile

    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
End Function


