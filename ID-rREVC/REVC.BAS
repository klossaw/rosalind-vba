Attribute VB_Name = "Ä£¿é2"
Option Explicit

Sub °´Å¥3_Click()
    Dim seq, revC_seq, letter As String
    Dim i As Integer
    seq = Range("a1").Value
    revC_seq = StrReverse(seq)
    For i = 1 To Len(revC_seq) Step 1
        letter = Mid(revC_seq, i, 1)
        If letter Like "A" Then
            revC_seq = Replace(revC_seq, letter, "T")
        ElseIf letter Like "C" Then
            revC_seq = Replace(revC_seq, letter, "G")
        ElseIf letter Like "G" Then
            revC_seq = Replace(revC_seq, letter, "C")
        ElseIf letter Like "T" Then
            revC_seq = Replace(revC_seq, letter, "A")
        End If
        Next i
        MsgBox revC_seq
End Sub
