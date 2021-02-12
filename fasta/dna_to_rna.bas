Attribute VB_Name = "Ä£¿é2"
Option Explicit


Public Sub DNA_TO_RNA()
' get row number
Dim row_number As Integer
row_number = Worksheets("fasta").UsedRange.Rows.Count

' get name and sequence
Dim name As String
Dim sequence()
ReDim sequence(1 To row_number - 1)
name = Worksheets("fasta").Cells(1, 1)
Dim i As Integer
For i = 1 To row_number - 1
sequence(i) = Worksheets("fasta").Cells(i + 1, 1)
Next i

'DNA-TO-RNA
Dim rna()
Dim j As Integer
ReDim rna(1 To row_number - 1)
For i = 1 To row_number - 1
 For j = 1 To Len(sequence(i))
      Select Case Mid(sequence(i), j, 1)
      Case Is = "A"
      Mid(sequence(i), j, 1) = "U"
       Case Is = "T"
      Mid(sequence(i), j, 1) = "A"
      Case Is = "C"
      Mid(sequence(i), j, 1) = "G"
      Case Is = "G"
     Mid(sequence(i), j, 1) = "C"
      End Select
 Next j
 Next i
 
'output
Worksheets.Add.name = "DNA-TO-RNA"
Worksheets("DNA-TO-RNA").Cells(1, 1) = name
For i = 1 To row_number - 1
Worksheets("DNA-TO-RNA").Cells(i + 1, 1) = sequence(i)
Next i
End Sub
