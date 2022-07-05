Sub GetFileName()
'ExtractFileName From Filepath
Dim Rng As Range
Dim WorkRng As Range
Dim splitList As Variant
On Error Resume Next
xTitleId = "GetFilename"
Set WorkRng = Application.Selection
Set WorkRng = Application.InputBox("Range", xTitleId, WorkRng.Address, Type:=8)
For Each Rng In WorkRng
    splitList = VBA.Split(Rng.Value, "\")
    Rng.Value = splitList(UBound(splitList, 1))
Next
End Sub