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
Private Sub UserForm_Activate()

    gerente1.Caption = Cells(2, 2).Value
    gerente1.Value = True
    gerente2.Caption = Cells(3, 2).Value
    
    linha = 9
    id = Cells(linha, 1).Value
    titulo = Cells(linha, 15).Value
    Solic = id & " - " & titulo
    
    While id <> ""
        Lista.AddItem (Solic)
        linha = linha + 1
        id = Cells(linha, 1).Value
        titulo = Cells(linha, 15).Value
        Solic = id & " - " & titulo
    Wend
    
End Sub
Private Sub Avancar_Click()
    Solic = Lista.Value
    If Lista.Value = Null Then
        ok = MsgBox("Tem que selecionar uma solicitação né LERDÃO!", vbOKOnly + vbExclamation, "Atenção")
        GoTo fim
    End If
    Dim id As Integer
    
    id = Left(Solic, InStr(1, Solic, " ") - 1) + 8
    nome = Cells(id, 15).Value
    titulo = "Solicitação de compras - " & nome
    
    If Cells(id, 8).Value = "" And Cells(id, 9).Value = "" Then
        
        Valor = ""
    
    ElseIf Cells(id, 8).Value <> "" Then
        preco = Format(Cells(id, 8).Value, "#,###.00")
        Valor = "<b>Valor:</b> R$" & preco
    
    Else
        preco = Format(Cells(id, 9).Value, "#,###.00")
        Valor = "<b>Valor:</b> U$" & preco
        
    End If
         
    If gerente1.Value = True Then
        destinatario = Cells(2, 3).Value
        gerente = Cells(2, 2).Value
    Else
        destinatario = Cells(3, 3).Value
        gerente = Cells(2, 3).Value
    End If

    If Valor <> "" Then
    
    mensagem = _
        "<font size='4' face='calibri'>" & _
        gerente & ", <br><br>" & _
        "Segue Nº do esboço referente à solicitação de compra <b><font color=#0066cc>" & nome & "</font></b>:<br><br>" & _
        "<b>Nº da chave: </b>" & Cells(id, 5).Value & "<br>" & _
        "<b>Nº do esboço: </b>" & Cells(id, 4).Value & "<br>" & _
        Valor & "<br><br>" & _
        "Aguardando aprovação.<br><br>" & _
        "Grato,</font>"
    
    Else
    
    mensagem = _
        "<font size='4' face='calibri'>" & _
        gerente & ", <br><br>" & _
        "Segue Nº do esboço referente à solicitação de compra <b><font color=#0066cc>" & nome & "</font></b>:<br><br>" & _
        "<b>Nº da chave: </b>" & Cells(id, 5).Value & "<br>" & _
        "<b>Nº do esboço: </b>" & Cells(id, 4).Value & "<br><br>" & _
        "Aguardando aprovação.<br><br>" & _
        "Grato,</font>"
    
    End If
    
    
    Set OutlookApp = CreateObject("Outlook.Application")
    Set OutlookMail = OutlookApp.CreateItem(0)
    
    OutlookMail.display
   
    With OutlookMail
      
      .to = destinatario
      .CC = "joao.gross@metaltork.com.br"
      .Subject = titulo
      .htmlBody = mensagem & OutlookMail.htmlBody
     End With


fim:
        
End Sub

Private Sub UserForm_Terminate()
    ActiveSheet.Shapes.Range(Array("EMAIL")).Select
    Selection.ShapeRange.Shadow.Visible = msoFalse
    Range("A9").Select
End Sub
