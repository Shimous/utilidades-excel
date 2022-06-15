Attribute VB_Name = "ETM"
Sub VOLTA_CADASTRO()
Attribute VOLTA_CADASTRO.VB_ProcData.VB_Invoke_Func = " \n14"
    Cells.Select
    Selection.EntireRow.Hidden = False
    Sheets("CADASTRO").Select
    Sheets("ETM").Visible = False
    
End Sub

Sub CADASTRO_ETM2()
    UF_ETM002.Show
    
End Sub

Sub REVISA_ETM2()
    ETM2R.Show
End Sub
Sub teste()
    MsgBox ThisWorkbook.Name
End Sub
