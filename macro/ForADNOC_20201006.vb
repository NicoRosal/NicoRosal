Private Sub CommandButton1_Click()
Dim SourceRange3 As Range, cel As Range

On Error Resume Next

   Set SourceRange3 = Application.Selection
   Set SourceRange3 = Application.InputBox("Range:", "Selece Filenames: ", SourceRange3.Address, Type:=8)
   
   Err.Clear

On Error GoTo 0
     
    For Each cel In SourceRange3
        If InStr(1, cel.Value, "UTILITY FLOW DIAGRAM") > 0 Then
            cel.Offset(0, 4).Value = "Process"
            cel.Offset(0, 5).Value = "Utility Flow Diagram"
            
        ElseIf InStr(1, cel.Value, "PROCESS FLOW DIAGRAM") > 0 Then
            cel.Offset(0, 4).Value = "Process"
            cel.Offset(0, 5).Value = "Process Flow Diagram"
          
        ElseIf InStr(1, cel.Value, "AUXILIARY FLOW DIAGRAM") > 0 Then
            cel.Offset(0, 4).Value = "Process"
            cel.Offset(0, 5).Value = "Process Flow Diagram"
                
        ElseIf InStr(1, cel.Value, "P&ID") > 0 Then
            cel.Offset(0, 4).Value = "Process"
            cel.Offset(0, 5).Value = "P&ID"
                            
        ElseIf InStr(1, cel.Value, "LOOP") > 0 Then
            cel.Offset(0, 4).Value = "Instrumentation"
            cel.Offset(0, 5).Value = "Instrument Loop Drawings"
            
        ElseIf InStr(1, cel.Value, "ISOMETRIC") > 0 Then
            cel.Offset(0, 4).Value = "Piping"
            cel.Offset(0, 5).Value = "Fabrication Isometric Drawing"
            
        ElseIf InStr(1, cel.Value, "CABLE BLOCK DIAGRAM") > 0 Then
            cel.Offset(0, 4).Value = "Instrumentation"
            cel.Offset(0, 5).Value = "Cable Block Diagram"
            
        ElseIf InStr(1, cel.Value, "PIPELINE ALIGNMENT") > 0 Then
            cel.Offset(0, 4).Value = "Pipeline"
            cel.Offset(0, 5).Value = "Pipeline Alignment Sheets"
            
   End If
   
   Next cel
   
End Sub

Private Sub GenerateUnits_Click()
Dim Tags As Variant
Dim TagUnit As Variant

Dim SrchRng As Range, cel As Range

Dim SourceRange As Range

On Error Resume Next

   Set SourceRange = Application.Selection
   Set SourceRange = Application.InputBox("Range:", "Select Tags", SourceRange.Address, Type:=8)
   
Err.Clear

On Error GoTo 0

   SourceRange.Offset(0, 2) = _
        "=IFERROR(LEFT(RC[-2],FIND(""-"",RC[-2],1)-1),LEFT(RC[-2],2))"
        

Set SrchRng = SourceRange

For Each cel In SrchRng
    If InStr(1, cel.Offset(0, 2).Value, "10") > 0 Then
        cel.Offset(0, 2).Value = "10-"

    ElseIf InStr(1, cel.Offset(0, 2).Value, "11") > 0 Then
        cel.Offset(0, 2).Value = "11-"
        
    ElseIf InStr(1, cel.Offset(0, 2).Value, "12") > 0 Then
        cel.Offset(0, 2).Value = "12-"
        
    ElseIf InStr(1, cel.Offset(0, 2).Value, "13") > 0 Then
        cel.Offset(0, 2).Value = "13-"

    ElseIf InStr(1, cel.Offset(0, 2).Value, "14") > 0 Then
        cel.Offset(0, 2).Value = "14-"
        
    ElseIf InStr(1, cel.Offset(0, 2).Value, "15") > 0 Then
        cel.Offset(0, 2).Value = "15-"
               
    ElseIf InStr(1, cel.Offset(0, 2).Value, "16") > 0 Then
        cel.Offset(0, 2).Value = "16-"
                             
    ElseIf InStr(1, cel.Offset(0, 2).Value, "19") > 0 Then
        cel.Offset(0, 2).Value = "19-"
    
    ElseIf InStr(1, cel.Offset(0, 2).Value, "52") > 0 Then
        cel.Offset(0, 2).Value = "52-"
        
    ElseIf InStr(1, cel.Offset(0, 2).Value, "45") > 0 Then
        cel.Offset(0, 2).Value = "45-"
     
    End If
Next cel
        
    For Each Tags In SourceRange.Offset(0, 2)
        If Tags = "14-" Then
             Tags.Offset(0, 5).Value = "Unit-14"
        
        ElseIf Tags = "13-" Then
             Tags.Offset(0, 5).Value = "Unit-13"
        ElseIf Tags = "11-" Then
             Tags.Offset(0, 5).Value = "Unit-11"
        ElseIf Tags = "12-" Then
             Tags.Offset(0, 5).Value = "Unit-12"
        ElseIf Tags = "15-" Then
             Tags.Offset(0, 5).Value = "Unit-15"
        ElseIf Tags = "16-" Then
             Tags.Offset(0, 5).Value = "Unit-16"
        ElseIf Tags = "10-" Then
             Tags.Offset(0, 5).Value = "Unit-10"
        ElseIf Tags = "45-" Then
             Tags.Offset(0, 5).Value = "Unit-45"
        ElseIf Tags = "52-" Then
             Tags.Offset(0, 5).Value = "Unit-15"
        ElseIf Tags = "83" Then
             Tags.Offset(0, 5).Value = "Unit-15"
        ElseIf Tags = "SS" Then
             Tags.Offset(0, 5).Value = "Unit-15"
        ElseIf Tags = "19-" Then
             Tags.Offset(0, 5).Value = "Unit-19"
        ElseIf Tags = "84" Then
             Tags.Offset(0, 5).Value = "Unit-15"
        ElseIf Tags = "SS" Then
             Tags.Offset(0, 5).Value = "Unit-15"
        ElseIf Tags = "SS1" Then
             Tags.Offset(0, 5).Value = "Unit-15"
        ElseIf Tags = "SS2" Then
             Tags.Offset(0, 5).Value = "Unit-15"
        ElseIf Tags = "SS3" Then
             Tags.Offset(0, 5).Value = "Unit-15"
        ElseIf Tags = "SS4" Then
             Tags.Offset(0, 5).Value = "Unit-15"
        ElseIf Tags = "SS5" Then
             Tags.Offset(0, 5).Value = "Unit-15"
        ElseIf Tags = "41" Then
             Tags.Offset(0, 5).Value = "Unit-15"
        ElseIf Tags = "42" Then
             Tags.Offset(0, 5).Value = "Unit-15"
        ElseIf Tags = "43" Then
             Tags.Offset(0, 5).Value = "Unit-15"
        ElseIf Tags = "44" Then
             Tags.Offset(0, 5).Value = "Unit-15"
        ElseIf Tags = "452" Then
             Tags.Offset(0, 5).Value = "Unit-00"
        ElseIf Tags = "47" Then
             Tags.Offset(0, 5).Value = "Unit-00"
        ElseIf Tags = "48" Then
             Tags.Offset(0, 5).Value = "Unit-15"
        ElseIf Tags = "49" Then
             Tags.Offset(0, 5).Value = "Unit-15"
        ElseIf Tags = "900" Then
             Tags.Offset(0, 5).Value = "Unit-900"
        ElseIf Tags = "573" Then
             Tags.Offset(0, 5).Value = "Unit-573"
        ElseIf Tags = "574" Then
             Tags.Offset(0, 5).Value = "Unit-573"
        ElseIf Tags = "401" Then
             Tags.Offset(0, 5).Value = "Unit-573"
        ElseIf Tags = "605" Then
             Tags.Offset(0, 5).Value = "Unit-00"
        ElseIf Tags = "015" Then
             Tags.Offset(0, 5).Value = "Unit-15"
        ElseIf Tags = "059" Then
             Tags.Offset(0, 5).Value = "Unit-15"
        ElseIf Tags = "17" Then
             Tags.Offset(0, 5).Value = "Unit-15"
        ElseIf Tags = "99" Then
             Tags.Offset(0, 5).Value = "Unit-99"
        ElseIf Tags = "" Then
             Tags.Offset(0, 5).Value = ""
        ElseIf Tags = "50" Then
             Tags.Offset(0, 5).Value = "Non-BAB"
        ElseIf Tags = "51" Then
             Tags.Offset(0, 5).Value = "Non-BAB"
        ElseIf Tags = "904" Then
             Tags.Offset(0, 5).Value = "Unit-00"
        ElseIf Tags = "96" Then
             Tags.Offset(0, 5).Value = "Unit-00"
        ElseIf Tags = "BRC" Then
             Tags.Offset(0, 5).Value = "Unit-00"
        ElseIf Tags = "WH" Then
             Tags.Offset(0, 5).Value = "Unit-00"
        ElseIf Tags = "21" Then
             Tags.Offset(0, 5).Value = "Non-BAB"
        ElseIf Tags = "22" Then
             Tags.Offset(0, 5).Value = "Non-BAB"
        ElseIf Tags = "25" Then
             Tags.Offset(0, 5).Value = "Non-BAB"
        ElseIf Tags = "50" Then
             Tags.Offset(0, 5).Value = "Non-BAB"
        ElseIf Tags = "51" Then
             Tags.Offset(0, 5).Value = "Non-BAB"
        ElseIf Tags = "545" Then
             Tags.Offset(0, 5).Value = "Non-BAB"
        ElseIf Tags = "61" Then
             Tags.Offset(0, 5).Value = "Non-BAB"
        ElseIf Tags = "62" Then
             Tags.Offset(0, 5).Value = "Non-BAB"
        ElseIf Tags = "64" Then
             Tags.Offset(0, 5).Value = "Non-BAB"
        ElseIf Tags = "65" Then
             Tags.Offset(0, 5).Value = "Non-BAB"
        ElseIf Tags = "66" Then
             Tags.Offset(0, 5).Value = "Non-BAB"
        ElseIf Tags = "67" Then
             Tags.Offset(0, 5).Value = "Non-BAB"
        ElseIf Tags = "68" Then
             Tags.Offset(0, 5).Value = "Non-BAB"
        ElseIf Tags = "70" Then
             Tags.Offset(0, 5).Value = "Non-BAB"
        ElseIf Tags = "71" Then
             Tags.Offset(0, 5).Value = "Non-BAB"
        ElseIf Tags = "72" Then
             Tags.Offset(0, 5).Value = "Non-BAB"
        ElseIf Tags = "80" Then
             Tags.Offset(0, 5).Value = "Non-BAB"
        ElseIf Tags = "81" Then
             Tags.Offset(0, 5).Value = "Non-BAB"
        ElseIf Tags = "82" Then
             Tags.Offset(0, 5).Value = "Non-BAB"
        ElseIf Tags = "080" Then
             Tags.Offset(0, 5).Value = "Non-BAB"
        ElseIf Tags = "92" Then
             Tags.Offset(0, 5).Value = "Non-BAB"
        Else
            Tags.Offset(0, 5).Value = ""
    End If
    
    Next Tags
    
    For Each TagUnit In SourceRange.Offset(0, 7)
        If TagUnit = "Unit-10" Then
            TagUnit.Offset(0, -1).Value = "Process"
        ElseIf TagUnit = "Unit-11" Then
            TagUnit.Offset(0, -1).Value = "Process"
        ElseIf TagUnit = "Unit-12" Then
            TagUnit.Offset(0, -1).Value = "Process"
        ElseIf TagUnit = "Unit-13" Then
            TagUnit.Offset(0, -1).Value = "Process"
        ElseIf TagUnit = "Unit-14" Then
            TagUnit.Offset(0, -1).Value = "Process"
        ElseIf TagUnit = "Unit-15" Then
            TagUnit.Offset(0, -1).Value = "Utility"
        ElseIf TagUnit = "Unit-16" Then
            TagUnit.Offset(0, -1).Value = "Utility"
        ElseIf TagUnit = "Unit-19" Then
            TagUnit.Offset(0, -1).Value = "Process"
        ElseIf TagUnit = "Unit-00" Then
            TagUnit.Offset(0, -1).Value = "Common"
        ElseIf TagUnit = "Unit-99" Then
            TagUnit.Offset(0, -1).Value = "Common"
        ElseIf TagUnit = "Unit-45" Then
            TagUnit.Offset(0, -1).Value = "Pipelines"
        ElseIf TagUnit = "Unit-900" Then
            TagUnit.Offset(0, -1).Value = "Pipelines"
        ElseIf TagUnit = "Unit-573" Then
            TagUnit.Offset(0, -1).Value = "Process"

    End If
    
    Next TagUnit

    
    Range("C:C").ClearContents
    Range("D:D").ClearContents
    Range("E:E").ClearContents
    Range("F:F").ClearContents
End Sub

Private Sub GetFileFormat_Click()


Dim SourceRange2 As Range

On Error Resume Next
   
   Set SourceRange2 = Application.Selection
   Set SourceRange2 = Application.InputBox("Range:", "Selece Filenames: ", SourceRange2.Address, Type:=8)
   
   Err.Clear

On Error GoTo 0
   
   SourceRange2.Offset(0, 1).Value = "=RIGHT(RC[-1],4)"
   
   Application.ScreenUpdating = False
   
   For Each fileform In SourceRange2.Offset(0, 1)
   
    If fileform = ".pdf" Then
        fileform.Value = "PDF"
        
    ElseIf fileform = ".dwg" Then
        fileform.Value = "CAD"

    ElseIf fileform = ".dgn" Then
        fileform.Value = "CAD"
        
    ElseIf fileform = ".xls" Then
        fileform.Value = "XLS"
        
     ElseIf fileform = "xlsx" Then
        fileform.Value = "XLS"
                      
     ElseIf fileform = ".doc" Then
        fileform.Value = "DOC"
        
    Else
        fileform.Value = ""
        
    End If
    
    Next fileform
    
      Application.ScreenUpdating = True
    
     SourceRange2.Offset(0, 2).Value = "ADNOC Gas Processing"
     SourceRange2.Offset(0, 3).Value = "Pipeline Network"
   
      
End Sub

Private Sub Label2_Click()

End Sub
