Sub CombineColumns1()
'updateby Extendoffice
    Dim xRng As Range
    Dim i As Integer
    Dim xLastRow As Integer
    Dim xTxt As String
    On Error Resume Next
    xTxt = Application.ActiveWindow.RangeSelection.Address
    Set xRng = Application.InputBox("please select the data range", "Kutools for Excel", xTxt, , , , , 8)
    If xRng Is Nothing Then Exit Sub
    xLastRow = xRng.Columns(1).Rows.Count + 1
    For i = 2 To xRng.Columns.Count
        Range(xRng.Cells(1, i), xRng.Cells(xRng.Columns(i).Rows.Count, i)).Cut
        ActiveSheet.Paste Destination:=xRng.Cells(xLastRow, 1)
        xLastRow = xLastRow + xRng.Columns(i).Rows.Count
    Next
End Sub