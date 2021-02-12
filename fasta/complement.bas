Attribute VB_Name = "Ä£¿é3"
Option Explicit


Public Sub complement()
' get row_number
Dim rownumber As Integer
rownumber = Worksheets("fasta").UsedRange.Rows.Count

' get name and sequence
Dim name As String
Dim sequence()
ReDim sequence(rownumber)
name = Worksheets("fasta").Cells(1, 1)
Dim i As Integer
For i = 1 To rownumber - 1
sequence(i) = Worksheets("fasta").Cells(i + 1, 1)
Next i

'complement
Dim j As Integer
Dim seq As String
For i = 1 To rownumber - 1
 seq = sequence(i)
 For j = 1 To Len(seq)
 Select Case Mid(seq, j, 1)
 Case Is = "A"
 Mid(seq, j, 1) = "T"
 Case Is = "T"
 Mid(seq, j, 1) = "A"
 Case Is = "C"
 Mid(seq, j, 1) = "G"
 Case Is = "G"
 Mid(seq, j, 1) = "C"
 End Select
 Next j
 Next i
 
'output
Worksheets.Add.name = "Complement"
Worksheets("Complement").Cells(1, 1) = name
For i = 1 To rownumber - 1
Worksheets("Complement").Cells(i + 1, 1) = sequence(i)
Next i

End Sub
