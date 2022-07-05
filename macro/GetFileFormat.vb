Sub GetFileFormat()

Dim SourceRange As Range

   Set SourceRange = Application.Selection
   Set SourceRange = Application.InputBox("Range:", "Selece Filenames: ", SourceRange.Address, Type:=8)
   
   SourceRange.Offset(0, 1).Value = "=RIGHT(RC[-1],4)"
   SourceRange.Offset(0, 2).Value = "ADNOC Gas Processing"
   SourceRange.Offset(0, 3).Value = "Pipeline Network"
   
   For Each fileform In Range("B2:B10000")
   
    If fileform = ".pdf" Then
        fileform.Value = "PDF"
        
    ElseIf fileform = ".dwg" Then
        fileform.Value = "CAD"
        
    Else
        fileform.Value = ""
        
    End If
    
    Next fileform
    
End Sub

