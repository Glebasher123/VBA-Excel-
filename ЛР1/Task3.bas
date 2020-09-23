Attribute VB_Name = "Module1"
Option Explicit
Sub PR2()
Dim a As Integer, b As Integer, c As Integer    ' описание переменных
Dim y As Double
a = Val(InputBox("Введите А"))                          ' ввод а
b = Val(InputBox("Введите В"))                          ' ввод b
c = Val(InputBox("Введите C"))                          ' ввод с

y = (Sqr(a + b) + b ^ 2) / (a + b + c) ^ 3 * Tan(a)
 ' вычисление  значения выражения
 
Range("A1").Value = " Значение A = "
Range("B1").Value = a

Range("A2").Value = " Значение B = "
Range("B2").Value = b

Range("A3").Value = " Значение C = "
Range("B3").Value = c

Range("A5").Value = " Значение Y = "
Range("B5").Value = y

End Sub
