Attribute VB_Name = "HackLink"
Sub LINK()
Attribute LINK.VB_ProcData.VB_Invoke_Func = " \n14"
'
' LINK Macro
    i = 12
    valor = Cells(i, 3).Value
    ID = "'" & Cells(i, 1).Value & "'!A1"
    
    While valor <> ""
    
    ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:= _
        ID, TextToDisplay:=valor
        i = i + 1
        valor = Cells(i, 3).Value
        ID = "'" & Cells(i, 1).Value & "'!A1"
        Cells(i, 3).Select
    Wend
    
End Sub

Sub macro_link()
    i = 3
    valor = Cells(i, 1).Value
    
    While valor <> ""
        ActiveSheet.Shapes.Range(Array("Retangulo_padrao")).Select
        Selection.Copy
        Cells(i, 1).Select
        ActiveSheet.Paste
        Selection.ShapeRange.IncrementLeft 1.764724094
        Selection.ShapeRange.IncrementTop 1.764724094
        Selection.ShapeRange.Name = valor

        Selection.OnAction = "Botao_link_Projeto"
        
        i = i + 1
        valor = Cells(i, 1).Value
        
    Wend
    
End Sub
