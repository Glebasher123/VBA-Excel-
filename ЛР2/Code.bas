Attribute VB_Name = "Module1"
Option Explicit

Sub ������_1()
  Dim x As Integer
  Dim y As Double
  x = Val(InputBox("������� x"))
  If x > 0 Then y = Sqr(x)
  If x < 0 Then y = x ^ 2
  If x = 0 Then y = 5
  MsgBox ("y=" & y)
End Sub

Sub ������_2()
  Dim x As Double, y As Double, z As Double, Max As Double
  x = Val(InputBox("������� x"))
  y = Val(InputBox("������� y"))
  z = Val(InputBox("������� z"))
  If (x > y) And (x > z) Then Max = x
  If (y > x) And (y > z) Then Max = y
  If (z > x) And (z > y) Then Max = z
  MsgBox ("��������=" & Max)
End Sub

Sub ������_3()
Dim opr As Double
Dim prem As Double
opr = Val(InputBox("������� ����� ������"))
Select Case opr
    Case 0 To 9999
         prem = 0.08 * opr
    Case 10000 To 39999
         prem = 0.1 * opr
    Case Is >= 40000
         prem = 0.14 * opr
End Select
MsgBox ("������������=" & prem)
End Sub

Sub �������_1()
    Dim x As Integer, y As Integer, z As Integer
    Dim f1 As Double, f2 As Double, f3 As Double
    x = Range("A2").Value
    y = Range("A3").Value
    z = Range("A4").Value
    f1 = (Application.WorksheetFunction.Max(x ^ 2, y ^ 2, x * z) + x)
    f2 = (Application.WorksheetFunction.Min(x, y) ^ 2 - y)
    f3 = f1 / f2
    MsgBox ("��������� =" & f3)
    Range("A5").Value = "��������� = "
    Range("B5").Value = f3
End Sub

Sub �������_2()
    Dim a As Integer, b As Integer, c As Double
    Dim d As Double, x1 As Double, x2 As Double
    a = Val(InputBox("������� a"))
    
    If a = 0 Then
        MsgBox ("� ��� �������� ���������")
        Return
    End If
        
    b = Val(InputBox("������� b"))
    c = Val(InputBox("������� c"))
    
    d = (b ^ 2) - (4 * a * c)
    Select Case d
         Case Is < 0
            MsgBox ("��� ������")
        Case Is = 0
            x1 = (-b) / (2 * a)
            MsgBox ("���� ������=" & x1)
        Case Is > 0
            x1 = ((-b + Sqr(d)) / (2 * a))
            x2 = ((-b - Sqr(d)) / (2 * a))
            MsgBox ("��� ����� =")
            MsgBox ("������ ������=" & x1)
            MsgBox ("������ ������=" & x2)
     End Select
End Sub

Sub �������_3_1()
Dim a As Integer, b As Integer
Dim res As Double
    a = Val(InputBox("������� �������� a"))
    b = Val(InputBox("������� �������� b"))
    
        If a > b Then
        MsgBox ("��������� = " & a * 2)
        Else
        MsgBox ("��������� = " & b * 2)
        End If
End Sub

Sub �������_3_2()
Dim x As Integer, y As Integer
  x = Val(InputBox("������� x"))
    If x > 0 Then
        y = 5 * x + x ^ 2
        MsgBox ("��������� = " & y)
        Range("F2").Value = "��������� = "
        Range("G2").Value = y
    ElseIf x <= 0 Then
        y = x + 2
        MsgBox ("��������� = " & y)
        Range("F2").Value = "��������� = "
        Range("G2").Value = y
    End If

End Sub


