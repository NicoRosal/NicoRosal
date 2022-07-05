Sub transposeColumns()
    Dim R1 As Range
    Dim R2 As Range
    Dim R3 As Range
    Dim RowN As Integer
    wTitle = "transpose multiple Columns"
    Set R1 = Application.Selection
    Set R1 = Application.InputBox("please select the Source data of Ranges:", wTitle, R1.Address, Type:=8)
    Set R2 = Application.InputBox("Select one destination single Cell or column:", wTitle, Type:=8)
    RowN = 0
    Application.ScreenUpdating = False
    For Each R3 In R1.Rows
        R3.Copy
        R2.Offset(RowN, 0).PasteSpecial Paste:=xlPasteAll, Transpose:=True
        RowN = RowN + R3.Columns.Count
    Next
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
End Sub
