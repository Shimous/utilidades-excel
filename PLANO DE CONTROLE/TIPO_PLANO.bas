Attribute VB_Name = "TIPO_PLANO"
Public wSheet As Worksheet

Sub PROT�TIPO()
Attribute PROT�TIPO.VB_ProcData.VB_Invoke_Func = " \n14"
'
' PROT�TIPO Macro
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
Sub PRODU��O()
Attribute PRODU��O.VB_ProcData.VB_Invoke_Func = " \n14"
'
' PRODU��O Macro
'

'
    
    With Sheets("MATERIA PRIMA")
        .Range("Prototype") = ""
        .Range("Prelaunch") = ""
        .Range("Production") = "X"
    End With

 
End Sub
