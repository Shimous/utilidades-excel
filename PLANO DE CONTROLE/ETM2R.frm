VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ETM2R 
   Caption         =   "Revisão ETM 002"
   ClientHeight    =   3690
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5175
   OleObjectBlob   =   "ETM2R.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ETM2R"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    sPos = ETM2R.ListBox1.Value
    Sheets("ETM").Visible = True
    Sheets("ETM").Select
    
    i = 0
    Linha = 2
    Pos = Cells(Linha, 4).Value
    
    While Pos <> ""
        If Pos = sPos Then
            If i = 0 Then GoTo Fim
            LinIni = Linha - 1
            Range(Cells(2, 4), Cells(LinIni, 4)).Select
            Selection.EntireRow.Hidden = True
Fim:
            While Pos = sPos
                Linha = Linha + 1
                Pos = Cells(Linha, 4).Value
            Wend
            LinFim = Linha
        End If
        i = 1
        Linha = Linha + 1
        Pos = Cells(Linha, 4).Value
    Wend
    
    Range(Cells(LinFim, 4), Cells(Linha, 4)).Select
    Selection.EntireRow.Hidden = True
    
    Cells(LinIni + 1, 4).Select
    
    ETM2R.ListBox1.Clear
    ETM2R.Hide
End Sub

Private Sub UserForm_Activate()

    Sheets("DADOS").Visible = True
    Sheets("DADOS").Select
    Linha = 2
    Pos = Cells(Linha, 5).Value
    While Pos <> ""
        ETM2R.ListBox1.AddItem (Pos)
        Linha = Linha + 1
        Pos = Cells(Linha, 5).Value
    Wend
    Sheets("DADOS").Visible = False
End Sub
