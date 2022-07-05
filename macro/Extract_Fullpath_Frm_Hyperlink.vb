Sub Extracthyperlinks()
'Extract Hyperlink chu2 (c) KutoolsforExcel
Dim Rng As Range
Dim WorkRng As Range
On Error Resume Next
xTitleId = "Hyperlink_Extractor_Chu2"
Set WorkRng = Application.Selection
Set WorkRng = Application.InputBox("Range", xTitleId, WorkRng.Address, Type:=8)
For Each Rng In WorkRng
    If Rng.Hyperlinks.Count > 0 Then
        Rng.Value = Rng.Hyperlinks.Item(1).Address
    End If
Next
End Sub