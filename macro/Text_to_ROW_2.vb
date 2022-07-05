Option Explicit

Sub Main()
'this is used for inserting into rows the tags with corresponding docnumber
    Columns("B:B").NumberFormat = "@"
    Dim i As Long, c As Long, r As Range, v As Variant

    For i = 1 To Range("B" & Rows.Count).End(xlUp).Row
        v = Split(Range("B" & i), ",")
        c = c + UBound(v) + 1
    Next i

    For i = 2 To c
        Set r = Range("B" & i)
        Dim arr As Variant
        arr = Split(r, ",")
        Dim j As Long
        r = arr(0)
        For j = 1 To UBound(arr)
            Rows(r.Row + j & ":" & r.Row + j).Insert Shift:=xlDown
            r.Offset(j, 0) = arr(j)
            r.Offset(j, -1) = r.Offset(0, -1)
            r.Offset(j, 1) = r.Offset(0, 1)
        Next j
    Next i

End Sub
