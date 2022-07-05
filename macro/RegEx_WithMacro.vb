Private Sub simpleRegex()
    Dim strPattern As String
    Dim strReplace As String: strReplace = ""
    Dim regEx As New RegExp
    Dim strInput As String
    Dim SourceRange As Range

    On Error Resume Next

       Set SourceRange = Application.Selection
       Set SourceRange = Application.InputBox("Range:", "Select Tags", SourceRange.Address, Type:=8)
       
       strPattern = Application.InputBox("Enter Pattern", Default:="800-*")
       
Err.Clear

On Error GoTo 0
    For Each cell In SourceRange
        If strPattern <> "" Then
            strInput = cell.Value

            With regEx
                .Global = True
                .MultiLine = True
                .IgnoreCase = False
                .Pattern = strPattern
            End With

            If regEx.Test(strInput) Then
                cell.Offset(0, 1).Value = "Match"
                            
            Else
                cell.Offset(0, 1).Value = "Not_Match"
            End If
        End If
    Next

    ActiveSheet.Range("A:C").AutoFilter Field:=3, Criteria1:="Match"
    Range("A:C").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveSheet.Previous.Select
    Selection.End(xlUp).Select
    Range("A:C").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Application.CutCopyMode = False
    Selection.EntireRow.Delete
    ActiveWindow.SmallScroll Down:=-9
    ActiveSheet.Range("A:C").AutoFilter Field:=3
    ActiveWindow.SmallScroll Down:=-3


End Sub
