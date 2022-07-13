VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CAD_PROJ 
   Caption         =   "Cadastro de Projeto"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5640
   OleObjectBlob   =   "CAD_PROJ.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CAD_PROJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cadastra_Click()
    Dim plan As String
    equipe = CAD_PROJ.equipe.Value
    titulo = CAD_PROJ.titulo.Value
    responsavel = CAD_PROJ.responsavel.Value
    
    If titulo = "" Then
        MsgBox "Digite um título para o projeto.", vbOKOnly + vbInformation, "Atenção"
        GoTo fim
    End If
    
    linha = AchaLinha("ID", 1)
    coluna = 1
    valor = Cells(linha, coluna).Value
    
    While valor <> ""
        linha = linha + 1
        valor = Cells(linha, coluna).Value
    
    Wend
 
    ActiveSheet.Unprotect
    ActiveWorkbook.Unprotect
    
    plan = linha - AchaLinha("ID", 1)
    
    Cells(linha, coluna).Value = plan
    c = Cells(linha, coluna).Value

'insere o titulo do projeto
    coluna = coluna + 2
    Cells(linha, coluna).Value = titulo

'insere o responsavel pelo projeto
    coluna = coluna + 1
    Cells(linha, coluna).Value = responsavel
    
'insere equipe do projeto
    coluna = coluna + 1
    Cells(linha, coluna).Value = equipe
    
    Sheets("MODELO").Visible = True
    Sheets("MODELO").Copy Before:=Sheets("Projetos")
    
    ThisWorkbook.ActiveSheet.Name = plan
    Range("A3:G3").Value = plan
    Range("A5").Select
    Sheets("MODELO").Visible = False
    
    Sheets("Projetos").Select
    
    ActiveSheet.Shapes.Range(Array("Retangulo_padrao")).Select
    Selection.Copy
    Cells(linha, 1).Select
    ActiveSheet.Paste
    
    With Selection.ShapeRange
        .IncrementLeft 1.764724094
        .IncrementTop 1.764724094
        .Name = plan
    End With
    
    Selection.OnAction = "BotaoLinkProjeto"
    
    'BLOCK
    
    Sheets(plan).Select
    Sheets("Projetos").Visible = False
    
    MkDir ("\\192.168.1.19\Dados\Arquivos\PROJETOS\1. CONTROLE DE PROJETOS\PROJETOS\" + plan + ". " + titulo)
    MkDir ("\\192.168.1.19\Dados\Arquivos\PROJETOS\1. CONTROLE DE PROJETOS\PROJETOS\" + plan + ". " + titulo + "\CUSTOS")
    MkDir ("\\192.168.1.19\Dados\Arquivos\PROJETOS\1. CONTROLE DE PROJETOS\PROJETOS\" + plan + ". " + titulo + "\DESENVOLVIMENTO")
   
Reset:
    With CAD_PROJ
        .responsavel.Value = ""
        .titulo.Value = ""
        .equipe.Value = ""
        .Hide
   End With
  
fim:
End Sub

Public Function AchaLinha(texto As String, coluna As Integer) As Integer
'texto = texto procurado
'coluna = coluna onde está procurando o texto

    linha = 1
    cel = Cells(linha, coluna).Value
    While cel <> texto And linha < 100000
        linha = linha + 1
        cel = Cells(linha, coluna).Value
    Wend
    
    If linha = 100000 Then
        MsgBox "Não consegui encontrar o que você está procurando!", vbExclamation, "Atenção"
    End If
    
    AchaLinha = linha
    
End Function


