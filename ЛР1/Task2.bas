Attribute VB_Name = "Module4"
Sub Задание2()
Dim a As Integer, b As Integer, Default As Integer, Default_1 As Integer

msg = "Введите число A"
Title = "Введите число A"
Default = 1
x = InputBox(msg, Title, Default, 4, 2)
Range("A1").Value = x

msg_1 = "Введите число B"
Title_1 = "Введите число B"
Default_1 = 0
x_1 = InputBox(msg_1, Title_1, Default_1, 5, 10)
Range("A2").Value = x_1

a = Range("A1").Value
b = Range("A2").Value

Style_1 = vbAbortRetryIgnore + vbExclamation
q = a + b
letter_1 = ("Сумма = " & q)
MsgBox letter_1, Style_1

Style_2 = vbRetryCancel + vbInformation
w = a * b
letter_2 = ("Произведение = " & w)
MsgBox letter_2, Style_2

Style_3 = vbOKCancel + vbQuestion
e = a Mod b
letter_3 = ("Частное = " & e)
MsgBox letter_3, Style_3

Style_4 = vbOKOnly + vbCritical
r = a - b
letter_4 = ("Разность = " & r)
MsgBox letter_4, Style_4

End Sub
