Attribute VB_Name = "BLOCK"
Public Sub BLOQUEIO(Optional ByVal ProtSheet As Worksheet)
Attribute BLOQUEIO.VB_ProcData.VB_Invoke_Func = " \n14"
'
' BLOQUEIO Macro
'

'

    If ProtSheet Is Nothing Then
    
        Set ProtSheet = ActiveSheet
    
    End If
    
    ProtSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:=True _
        , AllowFormattingCells:=True, AllowFormattingRows:=True, _
        AllowInsertingRows:=True, AllowInsertingHyperlinks:=True, _
        AllowDeletingRows:=True
    ActiveSheet.EnableSelection = xlNoRestrictions
End Sub

