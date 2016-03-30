Private Sub CommandButton1_Click()
'заполняет элементы массива рандомными положительными значениями от 0 до 1000'
For i = 1 To 30
Cells(1, i) = Int((1000 * Rnd) + 0)
Next i
End Sub

Private Sub CommandButton2_Click()
'находит и выводит сумму элементов массива, кратных 13'
For i = 1 To 30
If Cells(1, i) Mod 13 = 0 Then
a = a + Cells(1, i)
End If
Next i
MsgBox (a)
End Sub

Private Sub CommandButton3_Click()
'закрывает форму'
UserForm1.Hide
End Sub