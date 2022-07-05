Sub MoveVisible()
'Move Visible Cells to Tag Index
    Dim xRg As Range
    Dim xCell As Range
    Dim I As Long
    Dim J As Long
    Dim K As Long
    I = Worksheets("tag_automation").UsedRange.Rows.Count
    J = Worksheets("Tag_Index").UsedRange.Rows.Count
    
    If J = 1 Then
       If Application.WorksheetFunction.CountA(Worksheets("Tag_Index").UsedRange) = 0 Then J = 0
    End If
    
    Set xRg = Worksheets("tag_automation").Range("D1:D" & I)
    On Error Resume Next
    Application.ScreenUpdating = False
    
    For K = 1 To xRg.Count
        If CStr(xRg(K).Value) = "Done" Then
            xRg(K).EntireRow.Copy Destination:=Worksheets("Tag_Index").Range("A" & J + 1)
            If CStr(xRg(K).Value) = "Done" Then
                K = K - 1
            End If
            J = J + 1
        End If
    Next
    
    Application.ScreenUpdating = True
End Sub

Private Sub ImportTags()
'if docnumber same, retrieve tags then call move visible
    Dim docnum As Variant, cel As Range

    On Error Resume Next

    docnum = Application.InputBox("Enter Document Number", "Import Tags", "Document Number")
       
Err.Clear

On Error GoTo 0

   Application.ScreenUpdating = False
        
    For Each cell In Worksheets("tag_automation").Range("A1:A50")
    
        If cell.Value = docnum Then
            cell.Offset(0, 3).Value = "Done"
        Else
            cell.Offset(0, 3).Value = ""
            
        End If
		Call MoveVisible()
    Next cell
    


End Sub












