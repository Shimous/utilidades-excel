Attribute VB_Name = "REMOVE_JAT_MAGN"
Sub REMOVE_MAGN()
'
' REMOVE_MAGN Macro
'

'
    Sheets("MAGNAFLUX").Select
    ActiveWindow.SelectedSheets.Visible = False
End Sub
Sub REMOVE_JAT()
'
' REMOVE_JAT Macro
'

'
    Sheets("JATEAR").Select
    ActiveWindow.SelectedSheets.Visible = False
End Sub
