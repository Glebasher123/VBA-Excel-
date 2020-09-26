Sub Âàðèàíò5()
Dim x As Integer, y As Integer


x = Val(InputBox("Ââåäèòå ÷èñëî x"))
y = Val(InputBox("Ââåäèòå ÷èñëî y"))

Range("A1").Value = " Çíà÷åíèå X = "
Range("B1").Value = x

Range("A2").Value = " Çíà÷åíèå Y = "
Range("B2").Value = y

Z = ((y + 1) ^ 2) / 1 - ((x ^ 2) / (2 * y))
a = (Math.Sqr(x + y + Z) + Z ^ 2) / (1 + (y / 2) + (Z / 2))
b = x * y * Z * a ^ x - Math.Sin(a)
t = Math.Log(Math.Abs(x)) + Math.Exp(y)


Range("C1").Value = "Óðàâíåíèå Z = "
Range("D1").Value = Z

Range("C2").Value = "Óðàâíåíèå A = "
Range("D2").Value = a

Range("C3").Value = "Óðàâíåíèå B = "
Range("D3").Value = b

Range("C4").Value = "Óðàâíåíèå T = "
Range("D4").Value = t

End Sub

