Attribute VB_Name = "AchaLinha"
Public Function AchaLinha(ByVal texto As String, ByVal coluna As Integer, _
                            Optional ocorrencia As Integer = 1) As Integer
'texto = texto procurado
'coluna = coluna onde est� procurando o texto
'ocorrencia = padr�o primeira ocorrencia

        linha = 1
        cel = Cells(linha, coluna).Value
        x = 0
        
A:
        While cel <> texto And linha < 100000
            linha = linha + 1
            cel = Cells(linha, coluna).Value
        Wend
        
        x = x + 1
        If x <> ocorrencia Then
            linha = linha + 1
            cel = Cells(linha, coluna).Value
            GoTo A
            
        ElseIf linha = 100000 Then
            MsgBox "N�o consegui encontrar o que voc� est� procurando!", vbExclamation, "Aten��o"
            GoTo fim
        End If

    AchaLinha = linha
fim:
End Function

