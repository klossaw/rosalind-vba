Attribute VB_Name = "Ä£¿é1"
Option Explicit

Sub °´Å¥2_Click()
Dim month, offspring, result As Integer
month = InputBox("month:")
offspring = InputBox("offspring per month:")
Range("a2").Value = month
Range("b2").Value = offspring
result = fib(month, offspring)
MsgBox result
Range("a4").Value = result
End Sub

Public Function fib(ByVal n As Integer, ByVal k As Integer) As Variant
Dim fib0, fib1, sum, i As Integer
fib0 = 0
fib1 = 1
For i = 1 To n
sum = fib0 + k * fib1
fib0 = fib1
fib1 = sum
Next i
fib = fib0
End Function
