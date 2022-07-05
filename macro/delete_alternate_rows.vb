Sub Delete_Alternate_Rows_Excel()
    Dim SourceRange As Range
 
    Set SourceRange = Application.Selection
    Set SourceRange = Application.InputBox("Range:", "Select the range", SourceRange.Address, Type:=8)
 
    If SourceRange.Rows.Count >= 2 Then
        Dim FirstCell As Range
        Dim RowIndex As Integer
 
        Application.ScreenUpdating = False
 
        For RowIndex = SourceRange.Rows.Count To 1 Step -2
            Set FirstCell = SourceRange.Cells(RowIndex, 1)
            FirstCell.EntireRow.Delete
        Next
 
        Application.ScreenUpdating = True
 
    End If
End Sub