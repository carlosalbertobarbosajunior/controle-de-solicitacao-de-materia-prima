Attribute VB_Name = "M�dulo1"
Sub ultimaAtualizacao()
'
' Macro que atualiza o valor "�ltima atualiza��o" da c�lula H3

    Range("G6:I6").Value = Date & " �s " & Time()
    Range("B3").Select

End Sub

Sub Somar1()
' Aumenta em 1 o valor da c�lula de busca de servi�o
    Range("B3").Value = Range("B3").Value + 1

End Sub

Sub Subtrair1()
' Diminui em 1 o valor da c�lula de busca de servi�o
    Range("B3").Value = Range("B3").Value - 1
    
End Sub

