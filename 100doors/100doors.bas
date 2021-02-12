Attribute VB_Name = "Ä£¿é1"
Option Explicit

Sub genarate_Click()
' 100 doors from rosetta code
Dim initial(1 To 100) As integr
Dim i As Integer
For i = 1 To 100 Step 1
initial(i) = 0
Next i
Dim n As Integer
For n = 1 To 100
 For i = n To 100 Step n
 If initial(i) = 0 Then
 initial(i) = 1
 ElseIf initial(i) = 1 Then
 initial(i) = 0
 End If
 Next i
Next n
n = 1
Dim msgstr
For i = 1 To 100
If initial(i) = 1 Then
msgstr = msgstr & "   " & i
End If
Next i
MsgBox msgstr
End Sub
