Attribute VB_Name = "Module1"
Option Explicit
Sub PR2()
Dim a As Integer, b As Integer, c As Integer    ' �������� ����������
Dim y As Double
a = Val(InputBox("������� �"))                          ' ���� �
b = Val(InputBox("������� �"))                          ' ���� b
c = Val(InputBox("������� C"))                          ' ���� �

y = (Sqr(a + b) + b ^ 2) / (a + b + c) ^ 3 * Tan(a)
 ' ����������  �������� ���������
 
Range("A1").Value = " �������� A = "
Range("B1").Value = a

Range("A2").Value = " �������� B = "
Range("B2").Value = b

Range("A3").Value = " �������� C = "
Range("B3").Value = c

Range("A5").Value = " �������� Y = "
Range("B5").Value = y

End Sub
