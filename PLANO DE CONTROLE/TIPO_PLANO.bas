Attribute VB_Name = "TIPO_PLANO"
Public wSheet As Worksheet

Sub PROTÓTIPO()
Attribute PROTÓTIPO.VB_ProcData.VB_Invoke_Func = " \n14"
'
' PROTÓTIPO Macro
'

'
   
    With Sheets("MATERIA PRIMA")
        .Range("Prelaunch") = ""
        .Range("Production") = ""
        .Range("Prototype") = "X"
    End With
      
End Sub
Sub PRE_PROJETO()
Attribute PRE_PROJETO.VB_ProcData.VB_Invoke_Func = " \n14"
'
' PRE_PROJETO Macro
'

'
    With Sheets("MATERIA PRIMA")
        .Range("Prototype") = ""
        .Range("Production") = ""
        .Range("Prelaunch") = "X"
    End With
    
    
End Sub
Sub PRODUÇÃO()
Attribute PRODUÇÃO.VB_ProcData.VB_Invoke_Func = " \n14"
'
' PRODUÇÃO Macro
'

'
    
    With Sheets("MATERIA PRIMA")
        .Range("Prototype") = ""
        .Range("Prelaunch") = ""
        .Range("Production") = "X"
    End With

 
End Sub
