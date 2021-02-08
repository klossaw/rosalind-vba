Attribute VB_Name = "Ä£¿é1"
Option Explicit

Sub °´Å¥1_Click()
Dim seq1, seq2 As String
Dim hmd As Long, i As Long
seq1 = Range("a1").Value
seq2 = Range("a2").Value
If Len(seq1) <> Len(seq2) Then
MsgBox "sequences are of different length!"
End If
For i = 1 To Len(seq1)
If Mid(seq1, i, 1) <> Mid(seq2, i, 1) Then
hmd = hmd + 1
End If
Next i
MsgBox hmd
End Sub
