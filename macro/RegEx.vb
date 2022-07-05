Private Sub RegexMADAFAKA()
    Dim strPattern As String
    Dim strReplace As String: strReplace = ""
    Dim regEx As New RegExp
    Dim strInput As String
    Dim SourceRange As Range

    On Error Resume Next

       Set SourceRange = Application.Selection
       Set SourceRange = Application.InputBox("Range:", "Select Tags", SourceRange.Address, Type:=8)
       
       strPattern = Application.InputBox("Enter Pattern", Default:="[A-Z][A-Z]-[0-9][0-9][0-9]")
       
Err.Clear

On Error GoTo 0

   Application.ScreenUpdating = False
   
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
                cell.Offset(0, 1) = regEx.Execute(strInput)(0)
                            
            Else
                cell.Offset(0, 1).Value = ""
            End If
        End If
    Next

Application.ScreenUpdating = True


End Sub



