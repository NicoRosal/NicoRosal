Sub MoveVisible()
'Move Visible Cells to Sheet2
    Dim xRg As Range
    Dim xCell As Range
    Dim I As Long
    Dim J As Long
    Dim K As Long
    I = Worksheets("Sheet1").UsedRange.Rows.Count
    J = Worksheets("Sheet2").UsedRange.Rows.Count
	
    If J = 1 Then
       If Application.WorksheetFunction.CountA(Worksheets("Sheet2").UsedRange) = 0 Then J = 0
    End If
	
    Set xRg = Worksheets("Sheet1").Range("D1:D" & I)
    On Error Resume Next
    Application.ScreenUpdating = False
	
    For K = 1 To xRg.Count
        If CStr(xRg(K).Value) = "Done" Then
            xRg(K).EntireRow.Copy Destination:=Worksheets("Sheet2").Range("A" & J + 1)
            xRg(K).EntireRow.Delete
            If CStr(xRg(K).Value) = "Done" Then
                K = K - 1
            End If
            J = J + 1
        End If
    Next
	
    Application.ScreenUpdating = True
End Sub

Private Sub RegexMADAFAKA()
    Dim strPattern As String
    Dim strReplace As String: strReplace = ""
    Dim regEx As New RegExp
    Dim strInput As String
    Dim SourceRange As Range, cel as Range

    On Error Resume Next

       Set SourceRange = Application.Selection
       Set SourceRange = Application.InputBox("Range:", "Select Tags", SourceRange.Address, Type:=8)
       
	       
Err.Clear

On Error GoTo 0

   Application.ScreenUpdating = False
		
	For each cel in Worksheets("PatternToCheck").Range("B1:B3")
		
		StrPattern = cel.Value
		
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
					cell.Offset(0, 2) = "Done"
								
				Else
					cell.Offset(0, 1).Value = ""
				End If
			End If
		Next
	Call MoveVisible()
	Next cel

End Sub



