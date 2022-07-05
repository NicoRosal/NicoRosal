Private Sub Tag_Category()
Dim SourceRange3 As Range, cel As Range

On Error Resume Next

   Set SourceRange3 = Application.Selection
   Set SourceRange3 = Application.InputBox("Range:", "Selece Filenames: ", SourceRange3.Address, Type:=8)
   
   Err.Clear

On Error GoTo 0

   Application.ScreenUpdating = False
    
    SourceRange3.Offset(0, 1).Value = "=UPPER(RC[-1])"
        
    For Each cel In SourceRange3.Offset(0, 1)
    
        If InStr(1, cel.Value, "MOV") > 0 Then
            cel.Offset(0, 3).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 4).Value = "VALVE"
			cel.Offset(0, 5).Value = "ELECTRIC MOTOR OPERATED VALVE"
            
        ElseIf InStr(1, cel.Value, "ROV") > 0 Then
            cel.Offset(0, 3).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 4).Value = "VALVE"
			cel.Offset(0, 5).Value = "ELECTRIC MOTOR OPERATED VALVE"
			
        End If
   
    Next cel
	
       SourceRange3.Offset(0, 1).ClearContents
       
       For Each DT In SourceRange3.Offset(0, 3)

    If DT = "INSTRUMENT AND CONTROL" Then
        DT.Offset(0, 6).Value = "Instrumentation"

    ElseIf DT = "MECHANICAL" Then
        DT.Offset(0, 6).Value = "Mechanical"
        
        
    End If
    
        Next DT
		
End Sub