Attribute VB_Name = "Ä£¿é1"
Option Explicit

Sub °´Å¥1_Click()
Dim length_of_seq As Integer
Dim count_a, count_g, count_c, count_t, i As Integer
Dim seq, letter, msgstr As String
seq = Range("a1").Value
length_of_seq = Len(seq)
For i = 1 To length_of_seq Step 1
  letter = Mid(seq, i, 1)
  If letter Like "A" Then
    count_a = count_a + 1
  ElseIf letter Like "C" Then
    count_c = count_c + 1
  ElseIf letter Like "T" Then
    count_t = count_t + 1
  Else
    count_g = count_g + 1
  End If
Next i
msgstr = "A:" & count_a & Chr(10)
msgstr = msgstr & "C:" & count_c & Chr(10)
msgstr = msgstr & "G:" & count_g & Chr(10)
msgstr = msgstr & "T:" & count_t & Chr(10)
msgstr = msgstr & "total number is :" & length_of_seq
MsgBox msgstr
End Sub
