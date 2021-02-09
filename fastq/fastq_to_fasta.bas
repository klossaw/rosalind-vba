Attribute VB_Name = "Ä£¿é2"
Option Explicit


Public Sub fastq_to_fasta()
'get max row number
Dim row_number As Long
row_number = Worksheets("fastq").UsedRange.Rows.Count


'get seq name and sequence
Dim n As Long, i As Long
Dim seq_name(), sequence()
ReDim seq_name(row_number), sequence(row_number)
n = 1
For i = 1 To row_number Step 4
seq_name(n) = Worksheets("fastq").Cells(i, 1)
n = n + 1
Next i
n = 1
For i = 2 To row_number Step 4
sequence(n) = Worksheets("fastq").Cells(i, 1)
n = n + 1
Next i
n = n - 1

'fastq_to_fasta
Dim regex
Set regex = CreateObject("vbscript.regexp")
With regex
  .Global = True
  .Pattern = "@"
  End With
For i = 1 To n
seq_name(i) = regex.Replace(seq_name(i), ">")
Next i

'output
Worksheets.Add.Name = "fasta"
For i = 1 To n
Worksheets("fasta").Cells(2 * i - 1, 1) = seq_name(i)
Worksheets("fasta").Cells(2 * i, 1) = sequence(i)
Next i
End Sub
