Attribute VB_Name = "CARIMBO"
Public LstSht As Worksheet
Sub GoToLast()
    LstSht.Activate
End Sub

Sub CARIMBO()
Attribute CARIMBO.VB_ProcData.VB_Invoke_Func = " \n14"
'
' CARIMBO Macro
On Error GoTo msg
    If pShapeExists("EIMES", ActiveSheet) Or pShapeExists("CC", ActiveSheet) Then
        If pShapeExists("EIMES", ActiveSheet) Then
           ActiveSheet.Shapes.Range(Array("EIMES")).Select
           Selection.Delete
        End If
            
        If pShapeExists("CC", ActiveSheet) Then
           ActiveSheet.Shapes.Range(Array("CC")).Select
           Selection.Delete
        End If
        Exit Sub
    
    End If
    
    ActiveSheet.Unprotect
    Sheets("DADOS").Visible = True
    Sheets("DADOS").Select
    ActiveSheet.Shapes.Range(Array("data")).Select
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = Date
    ActiveSheet.Shapes.Range(Array("CC")).Select
    ActiveSheet.Shapes.Range(Array("CC", "EIMES")).Select
    Selection.Copy

    
' Volta para última planilha ativa

    GoToLast
    Sheets("DADOS").Visible = False
    Range("A18").Select
    ActiveSheet.Paste
    BLOQUEIO
    Range("A14").Select
    Exit Sub
msg:
    MsgBox "Selecione apenas uma planilha!", vbExclamation, "Atenção!"
End Sub
Private Function pShapeExists(sName As String, _
                              oSheet As Excel.Worksheet) As Boolean
  Dim oShape As Excel.Shape
  
  On Error Resume Next
  Set oShape = oSheet.Shapes(sName)
  On Error GoTo 0
  
  If Not oShape Is Nothing Then pShapeExists = True
End Function
