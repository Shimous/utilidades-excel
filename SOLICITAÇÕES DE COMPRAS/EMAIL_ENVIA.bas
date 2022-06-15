Attribute VB_Name = "EMAIL_ENVIA"

Sub GERA_EMAIL()

    ActiveSheet.Shapes.Range(Array("EMAIL")).Select
    Selection.ShapeRange.Shadow.Type = msoShadow30
    Range("A9").Select
    EMAIL_CONFIG.Show
    
    
End Sub
