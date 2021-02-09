Attribute VB_Name = "Ä£¿é3"
Option Explicit


Public Sub count_gc_length()
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

'count length
Dim count_length()
ReDim count_length(n)
For i = 1 To n
count_length(i) = Len(sequence(i))
Next i

'count GC
Dim GC()
ReDim GC(n)
Dim j As Long
For i = 1 To n
For j = 1 To Len(sequence(i))
If Mid(sequence(i), j, 1) Like "G" Or Mid(sequence(i), j, 1) Like "C" Then
GC(i) = GC(i) + 1
End If
Next j
Next i

'output
Worksheets.Add.Name = "count_gc"
Worksheets("count_gc").Cells(1, 1) = "Seq name"
Worksheets("count_gc").Cells(1, 2) = "Length"
Worksheets("count_gc").Cells(1, 3) = "GC content"
For i = 1 To n
Worksheets("count_gc").Cells(i + 1, 1) = seq_name(i)
Worksheets("count_gc").Cells(i + 1, 2) = count_length(i)
Worksheets("count_gc").Cells(i + 1, 3) = GC(i)
Next i
End Sub
