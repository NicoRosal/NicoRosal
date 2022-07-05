Private Sub Get_Discipline()
Dim cel As Range
Dim SourceRange as Range

On Error Resume Next

    Set SourceRange = Application.Selection
    Set SourceRange = Application.InputBox("Range:", "Select Tags: ", SourceRange3.Address, Type:=8)

    Err.Clear

On Error GoTo 0

Application ScreenUpdating = False

    For Each cel In SourceRange.Offset(0,3)

        If InStr(1, cel.Value, "INSTRUMENT AND CONTROL") > 0 Then
            cel.Offset(0, 6) = "Instrumentation"

        ElseIf InStr(1, cel.Value, "MECHANICAL") > 0 Then
            cel.Offset(0, 6) = "Mechanical"

        ElseIf InStr(1, cel.Value, "PIPING AND PIPELINE") > 0 Then
            cel.Offset(0, 6) = "Piping"

        ElseIf InStr(1, cel.Value, "CIVIL AND STRUCTURE") > 0 Then
            cel.Offset(0, 6) = "Civil & Structural"

        ElseIf InStr(1, cel.Value, "MISCELLANEOUS") > 0 Then
            cel.Offset(0, 6) = "TeleCommunication"

        ElseIf InStr(1, cel.Value, "HVAC EQUIPMENT") > 0 Then
            cel.Offset(0, 6) = "HVAC"

        ElseIf InStr(1, cel.Value, "ELECTRICAL") > 0 Then
            cel.Offset(0, 6) = "HVAC"

        ElseIf InStr(1, cel.Value, "HSE/ FIRE FIGHTING") > 0 Then
            cel.Offset(0, 6) = "HSE"
            
    End If
    
Next cel

   For Each ST In SourceRange.Offset(0, 5)
       
    If ST = "PIPERUN" Then
        ST.Offset(0, 4).Value = "Pipeline"
        
    End If
    
        Next ST
        
   For Each TT In SourceRange.Offset(0, 2)
       
    If TT = "DELETE" Then
        TT.Offset(0, 1).Value = "DELETE"
        TT.Offset(0, 20).Value = "DELETE"
        
    End If
    
        Next TT
        
End Sub



            
