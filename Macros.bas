Attribute VB_Name = "Módulo1"
Sub ultimaAtualizacao()
'
' Macro que atualiza o valor "Última atualização" da célula H3

    Range("G6:I6").Value = Date & " às " & Time()
    Range("B3").Select

End Sub

Sub Somar1()
' Aumenta em 1 o valor da célula de busca de serviço
    Range("B3").Value = Range("B3").Value + 1

End Sub

Sub Subtrair1()
' Diminui em 1 o valor da célula de busca de serviço
    Range("B3").Value = Range("B3").Value - 1
    
End Sub

