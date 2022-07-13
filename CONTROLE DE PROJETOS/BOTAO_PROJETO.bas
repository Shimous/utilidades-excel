Attribute VB_Name = "BOTAO_PROJETO"
Sub BotaoLinkProjeto()
Attribute BotaoLinkProjeto.VB_ProcData.VB_Invoke_Func = " \n14"
    
    Sheets(Application.Caller).Visible = True
    Sheets("Projetos").Visible = False
    Sheets(Application.Caller).Select
    Range("A2").Select

End Sub
