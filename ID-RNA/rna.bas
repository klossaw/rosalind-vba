Attribute VB_Name = "ģ��1"
Option Explicit

Sub ��ť1_Click()
    Dim seq, letter As String
    Dim i As Integer
    seq = Range("a1").Value
    For i = 1 To Len(seq)
        letter = Mid(seq, i, 1)
        If letter Like "T" Then
            seq = Replace(seq, letter, "U")
        End If
    Next i
    MsgBox seq
End Sub
