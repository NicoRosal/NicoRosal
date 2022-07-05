Sub ConcatenateCellsIfSameValues()
'Concatenating Reference Doc/Tag to filename

    Dim xCol As New Collection
    Dim xSrc As Variant
    Dim xRes() As Variant
    Dim I As Long
    Dim J As Long
    Dim xRg As Range
    xSrc = Range("A1", Cells(Rows.Count, "A").End(xlUp)).Resize(, 2)
    Set xRg = Range("D1")
    On Error Resume Next
    For I = 2 To UBound(xSrc)
        xCol.Add xSrc(I, 1), TypeName(xSrc(I, 1)) & CStr(xSrc(I, 1))
    Next I
    On Error GoTo 0
    ReDim xRes(1 To xCol.Count + 1, 1 To 2)
    xRes(1, 1) = "Filename"
    xRes(1, 2) = "Reference Documents/Tags"
    For I = 1 To xCol.Count
        xRes(I + 1, 1) = xCol(I)
        For J = 2 To UBound(xSrc)
            If xSrc(J, 1) = xRes(I + 1, 1) Then
                xRes(I + 1, 2) = xRes(I + 1, 2) & ", " & xSrc(J, 2)
            End If
        Next J
        xRes(I + 1, 2) = Mid(xRes(I + 1, 2), 2)
    Next I
    Set xRg = xRg.Resize(UBound(xRes, 1), UBound(xRes, 2))
    xRg.NumberFormat = "@"
    xRg = xRes
    xRg.EntireColumn.AutoFit
End Sub
