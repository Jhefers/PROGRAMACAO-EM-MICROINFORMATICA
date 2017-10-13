Sub variacao()
Dim n As Integer, i As Integer
Dim maior As Integer, menor As Integer

For i = 1 To 10
    n = InputBox("Digite um valor: ")
    If i = 1 Then
        maior = n
        menor = n
    End If
    If n > maior Then maior = n
    If n < menor Then menor = n
Next
MsgBox "Maior valor: " & maior
MsgBox "Menor valor: " & menor

    
End Sub
