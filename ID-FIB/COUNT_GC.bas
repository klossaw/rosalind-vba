Attribute VB_Name = "Ä£¿é1"
Option Explicit

Sub °´Å¥1_Click()
'get max townumber
Dim row_number As Long
row_number = Worksheets("sheet1").UsedRange.Rows.Count


'get the number of sequence name
Dim regex As Object
Set regex = CreateObject("vbscript.regexp")
  With regex
    .Global = True
    .pattern = ">"
    End With
Dim num_seqname As Long, i As Long
num_seqname = 0
For i = 1 To row_number
If regex.test(Cells(i, 1)) = True Then
num_seqname = num_seqname + 1
End If
Next i


'get the sequence name and the location
Dim seq_name() As String
Dim name_number() As Long
Dim n As Long
n = 1
ReDim seq_name(num_seqname), name_number(num_seqname)
For i = 1 To row_number
If regex.test(Cells(i, 1)) = True Then
seq_name(n) = Cells(i, 1)
name_number(n) = i
n = n + 1
End If
Next i

'compute content of GC
Dim begin, finish As Long
Dim GC() As Long
Dim total() As Long
ReDim total(num_seqname)
ReDim GC(num_seqname)
Dim count_g, count_c As Long
Dim j As Long, k As Long
For j = 1 To (num_seqname - 1)
begin = name_number(j) + 1
finish = name_number(j + 1) - 1
For i = begin To finish
For k = 1 To Len(Cells(i, 1))
If Mid(Cells(i, 1), k, 1) Like "G" Then
count_g = count_g + 1
ElseIf Mid(Cells(i, 1), k, 1) Like "C" Then
count_c = count_c + 1
End If
Next k
total(j) = total(j) + Len(Cells(i, 1))
Next i
GC(j) = count_g + count_c
count_g = 0
count_c = 0
Next j
begin = name_number(num_seqname)
finish = row_number
For i = begin To finish
For k = 1 To Len(Cells(i, 1))
If Mid(Cells(i, 1), k, 1) Like "G" Then
count_g = count_g + 1
ElseIf Mid(Cells(i, 1), k, 1) Like "C" Then
count_c = count_c + 1
End If
Next k
total(num_seqname) = total(num_seqname) + Len(Cells(i, 1))
Next i
GC(num_seqname) = count_g + count_c
Dim mgstr As String
For i = 1 To num_seqname
mgstr = mgstr & seq_name(i) & ":" & "total:" & total(i) & "   " & "content of GC:" & GC(i) & Chr(10)
Next i
MsgBox mgstr
End Sub
