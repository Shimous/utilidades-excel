Attribute VB_Name = "AchaLinha"
Public Function AchaLinha(ByVal texto As String, ByVal coluna As Integer, _
                            Optional ocorrencia As Integer = 1) As Integer
'texto = texto procurado
'coluna = coluna onde está procurando o texto
'ocorrencia = padrão primeira ocorrencia

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
            MsgBox "Não consegui encontrar o que você está procurando!", vbExclamation, "Atenção"
            GoTo fim
        End If

    AchaLinha = linha
fim:
End Function

