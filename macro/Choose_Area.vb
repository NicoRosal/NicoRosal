Private Sub GetUnit()
Dim SourceRange1 As Range
Dim cel As Range

On Error Resume Next

    Set SourceRange1 = Application.Selection
    Set SourceRange1 = Application.InputBox("Range:", "Select Units: ", SourceRange1.Address, Type:=8)


   
   
   Err.Clear
   
   On Error GoTo 0
   
      Application.ScreenUpdating = False
        
    For Each cel In SourceRange1
            
         If InStr(1, cel.Value, "511A") = 1 Then
            cel.Offset(0, -1).Value = "Maqta"
            
         ElseIf InStr(1, cel.Value, "11-01") = 1 Then
            cel.Offset(0, -1).Value = "Bab"
            
         ElseIf InStr(1, cel.Value, "4243") = 1 Then
            cel.Offset(0, -1).Value = "Ruwais"
            
         ElseIf InStr(1, cel.Value, "997") = 1 Then
            cel.Offset(0, -1).Value = "Al Ain"
            
         ElseIf InStr(1, cel.Value, "993") = 1 Then
            cel.Offset(0, -1).Value = "Maqta"
            
         ElseIf InStr(1, cel.Value, "983") = 1 Then
            cel.Offset(0, -1).Value = "Mirfa-Ruwais"
            
         ElseIf InStr(1, cel.Value, "981") = 1 Then
            cel.Offset(0, -1).Value = "Habshan-Mirfa"
            
         ElseIf InStr(1, cel.Value, "950") = 1 Then
            cel.Offset(0, -1).Value = "Habshan-Ruwais"
            
         ElseIf InStr(1, cel.Value, "949") = 1 Then
            cel.Offset(0, -1).Value = "Habshan-Ruwais"
            
        ElseIf InStr(1, cel.Value, "948") = 1 Then
            cel.Offset(0, -1).Value = "Al Ain"
        
        ElseIf InStr(1, cel.Value, "947") = 1 Then
            cel.Offset(0, -1).Value = "Habshan"
        
        ElseIf InStr(1, cel.Value, "944") = 1 Then
            cel.Offset(0, -1).Value = "Habshan-Bab"
            
        ElseIf InStr(1, cel.Value, "943") = 1 Then
            cel.Offset(0, -1).Value = "Habshan"
        
        ElseIf InStr(1, cel.Value, "942") = 1 Then
            cel.Offset(0, -1).Value = "Bab-Ruwais"
        
        ElseIf InStr(1, cel.Value, "941") = 1 Then
        cel.Offset(0, -1).Value = "Habshan-Ruwais"
        
        ElseIf InStr(1, cel.Value, "931") = 1 Then
        cel.Offset(0, -1).Value = "Bu Hasa-MP21"
        
        ElseIf InStr(1, cel.Value, "922") = 1 Then
        cel.Offset(0, -1).Value = "Asab-Bab"
        
        ElseIf InStr(1, cel.Value, "921") = 1 Then
        cel.Offset(0, -1).Value = "Asab-Habshan"
        
        ElseIf InStr(1, cel.Value, "903") = 1 Then
        cel.Offset(0, -1).Value = "Habshan-Bab"
        
        ElseIf InStr(1, cel.Value, "902") = 1 Then
        cel.Offset(0, -1).Value = "Ruwais"
        
        ElseIf InStr(1, cel.Value, "901") = 1 Then
        cel.Offset(0, -1).Value = "Habshan-Ruwais"
        
        ElseIf InStr(1, cel.Value, "900") = 1 Then
        cel.Offset(0, -1).Value = "Bab"
        
        ElseIf InStr(1, cel.Value, "892") = 1 Then
        cel.Offset(0, -1).Value = "Musaffah"
        
        ElseIf InStr(1, cel.Value, "887") = 1 Then
        cel.Offset(0, -1).Value = "Maqta-Jebel Ali"
        
        ElseIf InStr(1, cel.Value, "834") = 1 Then
        cel.Offset(0, -1).Value = "Maqta-Taweelah"
        
        ElseIf InStr(1, cel.Value, "830") = 1 Then
        cel.Offset(0, -1).Value = "Shahama-Mina Zayed"
        
        ElseIf InStr(1, cel.Value, "827") = 1 Then
        cel.Offset(0, -1).Value = "Maqta-Taweelah"
        
        ElseIf InStr(1, cel.Value, "824") = 1 Then
        cel.Offset(0, -1).Value = "Al Ain"
        
        ElseIf InStr(1, cel.Value, "823") = 1 Then
        cel.Offset(0, -1).Value = "Musaffah"
        
        ElseIf InStr(1, cel.Value, "821") = 1 Then
        cel.Offset(0, -1).Value = "Shahama-Mina Zayed"
        
        ElseIf InStr(1, cel.Value, "819") = 1 Then
        cel.Offset(0, -1).Value = "Maqta-Taweelah"
        
        ElseIf InStr(1, cel.Value, "818") = 1 Then
        cel.Offset(0, -1).Value = "Musaffah"
        
        ElseIf InStr(1, cel.Value, "817") = 1 Then
        cel.Offset(0, -1).Value = "Maqta-Al Ain"
       
        ElseIf InStr(1, cel.Value, "816") = 1 Then
        cel.Offset(0, -1).Value = "Maqta-Al Ain"
        
        ElseIf InStr(1, cel.Value, "815") = 1 Then
        cel.Offset(0, -1).Value = "Maqta-Jebel Ali"
        
        ElseIf InStr(1, cel.Value, "814") = 1 Then
        cel.Offset(0, -1).Value = "Maqta-Jebel Ali"
        
        ElseIf InStr(1, cel.Value, "813") = 1 Then
        cel.Offset(0, -1).Value = "Musaffah"
        
        ElseIf InStr(1, cel.Value, "812") = 1 Then
        cel.Offset(0, -1).Value = "Maqta-Al Ain"
        
        ElseIf InStr(1, cel.Value, "811") = 1 Then
        cel.Offset(0, -1).Value = "Maqta-Jebel Ali"
        
        ElseIf InStr(1, cel.Value, "809") = 1 Then
        cel.Offset(0, -1).Value = "Al Ain"
        
        ElseIf InStr(1, cel.Value, "808") = 1 Then
        cel.Offset(0, -1).Value = "Al Ain"
        
        ElseIf InStr(1, cel.Value, "807") = 1 Then
        cel.Offset(0, -1).Value = "Maqta-Taweelah"
        
        ElseIf InStr(1, cel.Value, "801") = 1 Then
        cel.Offset(0, -1).Value = "Abu Dhabi Island"
        
        ElseIf InStr(1, cel.Value, "800") = 1 Then
        cel.Offset(0, -1).Value = "Maqta"
        
        ElseIf InStr(1, cel.Value, "766") = 1 Then
        cel.Offset(0, -1).Value = "Maqta"
        
        ElseIf InStr(1, cel.Value, "714") = 1 Then
        cel.Offset(0, -1).Value = "Ruwais"
        
        ElseIf InStr(1, cel.Value, "713") = 1 Then
        cel.Offset(0, -1).Value = "Ruwais"
        
        ElseIf InStr(1, cel.Value, "712") = 1 Then
        cel.Offset(0, -1).Value = "Ruwais"
        
        ElseIf InStr(1, cel.Value, "711") = 1 Then
        cel.Offset(0, -1).Value = "Ruwais-Shuweihat"
        
        ElseIf InStr(1, cel.Value, "710") = 1 Then
        cel.Offset(0, -1).Value = "Ruwais-Shuweihat"
        
        ElseIf InStr(1, cel.Value, "709") = 1 Then
        cel.Offset(0, -1).Value = "Ruwais"
        
        ElseIf InStr(1, cel.Value, "708") = 1 Then
        cel.Offset(0, -1).Value = "Ruwais"
        
        ElseIf InStr(1, cel.Value, "706") = 1 Then
        cel.Offset(0, -1).Value = "Ruwais"
        
        ElseIf InStr(1, cel.Value, "705") = 1 Then
        cel.Offset(0, -1).Value = "Habshan-Mirfa"
        
        ElseIf InStr(1, cel.Value, "704") = 1 Then
        cel.Offset(0, -1).Value = "Ruwais"
        
        ElseIf InStr(1, cel.Value, "702") = 1 Then
        cel.Offset(0, -1).Value = "Ruwais"
        
        ElseIf InStr(1, cel.Value, "701") = 1 Then
        cel.Offset(0, -1).Value = "Ruwais"
        
        ElseIf InStr(1, cel.Value, "605") = 1 Then
        cel.Offset(0, -1).Value = "Madinat Zayed"
       
        ElseIf InStr(1, cel.Value, "603") = 1 Then
        cel.Offset(0, -1).Value = "Habshan"
        
        ElseIf InStr(1, cel.Value, "602") = 1 Then
        cel.Offset(0, -1).Value = "Madinat Zayed"
        
        ElseIf InStr(1, cel.Value, "601") = 1 Then
        cel.Offset(0, -1).Value = "Thamamma C"
        
        ElseIf InStr(1, cel.Value, "600") = 1 Then
        cel.Offset(0, -1).Value = "Bab"
        
        ElseIf InStr(1, cel.Value, "594") = 1 Then
        cel.Offset(0, -1).Value = "Shahama-Mina Zayed"
        
        ElseIf InStr(1, cel.Value, "592") = 1 Then
        cel.Offset(0, -1).Value = "Ras Al Qila-Habshan"
        
        ElseIf InStr(1, cel.Value, "590") = 1 Then
        cel.Offset(0, -1).Value = "Al Ain"
        
        ElseIf InStr(1, cel.Value, "588") = 1 Then
        cel.Offset(0, -1).Value = "Ruwais"
        
        ElseIf InStr(1, cel.Value, "586") = 1 Then
        cel.Offset(0, -1).Value = "Ruwais"
        
        ElseIf InStr(1, cel.Value, "585") = 1 Then
        cel.Offset(0, -1).Value = "Bab-Thamamma C"
        
        ElseIf InStr(1, cel.Value, "584") = 1 Then
        cel.Offset(0, -1).Value = "Bu Hasa-Bab"
        
        ElseIf InStr(1, cel.Value, "582") = 1 Then
        cel.Offset(0, -1).Value = "Thamamma C-Maqta"
        
        ElseIf InStr(1, cel.Value, "581") = 1 Then
        cel.Offset(0, -1).Value = "Shahama-Mina Zayed"
       
        ElseIf InStr(1, cel.Value, "578") = 1 Then
        cel.Offset(0, -1).Value = "Maqta-Taweelah"
        
        ElseIf InStr(1, cel.Value, "577") = 1 Then
        cel.Offset(0, -1).Value = "Thamamma C-Maqta"
        
        ElseIf InStr(1, cel.Value, "573") = 1 Then
        cel.Offset(0, -1).Value = "Bab"
        
        ElseIf InStr(1, cel.Value, "571") = 1 Then
        cel.Offset(0, -1).Value = "Ruwais-Shuweihat"
        
        ElseIf InStr(1, cel.Value, "570") = 1 Then
        cel.Offset(0, -1).Value = "Thamamma C-Ruwais"
        
        ElseIf InStr(1, cel.Value, "569") = 1 Then
        cel.Offset(0, -1).Value = "Madinat Zayed"
        
        ElseIf InStr(1, cel.Value, "568") = 1 Then
        cel.Offset(0, -1).Value = "Bab"
        
        ElseIf InStr(1, cel.Value, "566") = 1 Then
        cel.Offset(0, -1).Value = "Musaffah"
        
        ElseIf InStr(1, cel.Value, "564") = 1 Then
        cel.Offset(0, -1).Value = "Maqta-Jebel Ali"
        
        ElseIf InStr(1, cel.Value, "561") = 1 Then
        cel.Offset(0, -1).Value = "Asab"
        
        ElseIf InStr(1, cel.Value, "560") = 1 Then
        cel.Offset(0, -1).Value = "Maqta-Taweelah"
       
        ElseIf InStr(1, cel.Value, "557") = 1 Then
        cel.Offset(0, -1).Value = "Thamamma C-Asab"
        
        ElseIf InStr(1, cel.Value, "556") = 1 Then
        cel.Offset(0, -1).Value = "Asab"
        
        ElseIf InStr(1, cel.Value, "555") = 1 Then
        cel.Offset(0, -1).Value = "Asab"
        
        ElseIf InStr(1, cel.Value, "553") = 1 Then
        cel.Offset(0, -1).Value = "Ras Al Qila-Habshan"
        
        ElseIf InStr(1, cel.Value, "552") = 1 Then
        cel.Offset(0, -1).Value = "Bu Hasa-Habshan"
        
        ElseIf InStr(1, cel.Value, "551") = 1 Then
        cel.Offset(0, -1).Value = "Thamamma C-Ruwais"
        
        ElseIf InStr(1, cel.Value, "550") = 1 Then
        cel.Offset(0, -1).Value = "Thamamma C-Maqta"
        
        ElseIf InStr(1, cel.Value, "545") = 1 Then
        cel.Offset(0, -1).Value = "Habshan-Bab"
        
        ElseIf InStr(1, cel.Value, "541") = 1 Then
        cel.Offset(0, -1).Value = "Habshan"
        
        ElseIf InStr(1, cel.Value, "540") = 1 Then
        cel.Offset(0, -1).Value = "Habshan"
        
        ElseIf InStr(1, cel.Value, "520") = 1 Then
        cel.Offset(0, -1).Value = "Habshan-Maqta"
        
        ElseIf InStr(1, cel.Value, "519") = 1 Then
        cel.Offset(0, -1).Value = "Al Ain"
        
        ElseIf InStr(1, cel.Value, "518") = 1 Then
        cel.Offset(0, -1).Value = "Maqta-Al Ain"
        
        ElseIf InStr(1, cel.Value, "517") = 1 Then
        cel.Offset(0, -1).Value = "Musaffah"
        
        ElseIf InStr(1, cel.Value, "516") = 1 Then
        cel.Offset(0, -1).Value = "Maqta-Al Ain"
        
        ElseIf InStr(1, cel.Value, "515") = 1 Then
        cel.Offset(0, -1).Value = "Maqta-Jebel Ali"
        
        ElseIf InStr(1, cel.Value, "514") = 1 Then
        cel.Offset(0, -1).Value = "Maqta-Al Ain"
        
        ElseIf InStr(1, cel.Value, "512") = 1 Then
        cel.Offset(0, -1).Value = "Maqta-Taweelah"
        
        ElseIf InStr(1, cel.Value, "510") = 1 Then
        cel.Offset(0, -1).Value = "Maqta"
        
        ElseIf InStr(1, cel.Value, "508") = 1 Then
        cel.Offset(0, -1).Value = "Ruwais"
        
        ElseIf InStr(1, cel.Value, "506") = 1 Then
        cel.Offset(0, -1).Value = "Thamamma C-Mirfa"
       
        ElseIf InStr(1, cel.Value, "505") = 1 Then
        cel.Offset(0, -1).Value = "Bu Hasa-Bab"
        
        ElseIf InStr(1, cel.Value, "504") = 1 Then
        cel.Offset(0, -1).Value = "Asab"
        
        ElseIf InStr(1, cel.Value, "503") = 1 Then
        cel.Offset(0, -1).Value = "Bab-Ruwais"
        
        ElseIf InStr(1, cel.Value, "502") = 1 Then
        cel.Offset(0, -1).Value = "Thamamma C-Maqta"
        
        ElseIf InStr(1, cel.Value, "501") = 1 Then
        cel.Offset(0, -1).Value = "Bab-Maqta"
        
        ElseIf InStr(1, cel.Value, "403") = 1 Then
        cel.Offset(0, -1).Value = "Asab"
        
        ElseIf InStr(1, cel.Value, "402") = 1 Then
        cel.Offset(0, -1).Value = "Bu Hasa"
        
        ElseIf InStr(1, cel.Value, "401") = 1 Then
        cel.Offset(0, -1).Value = "Bab"
        
        ElseIf InStr(1, cel.Value, "377") = 1 Then
        cel.Offset(0, -1).Value = "Habshan"
        
        ElseIf InStr(1, cel.Value, "326") = 1 Then
        cel.Offset(0, -1).Value = "Habshan"
        
        ElseIf InStr(1, cel.Value, "273") = 1 Then
        cel.Offset(0, -1).Value = "Ras Al Qila-Habshan"
        
        ElseIf InStr(1, cel.Value, "203") = 1 Then
        cel.Offset(0, -1).Value = "Ruwais"
        
        ElseIf InStr(1, cel.Value, "202") = 1 Then
        cel.Offset(0, -1).Value = "Habshan"
        
        ElseIf InStr(1, cel.Value, "201") = 1 Then
        cel.Offset(0, -1).Value = "Al Ain"
        
        ElseIf InStr(1, cel.Value, "200") = 1 Then
        cel.Offset(0, -1).Value = "Asab"
        
        ElseIf InStr(1, cel.Value, "190") = 1 Then
        cel.Offset(0, -1).Value = "Asab"
        
        ElseIf InStr(1, cel.Value, "173") = 1 Then
        cel.Offset(0, -1).Value = "Bu Hasa-Habshan"
        
        ElseIf InStr(1, cel.Value, "127") = 1 Then
        cel.Offset(0, -1).Value = "Asab-Ruwais"
        
        ElseIf InStr(1, cel.Value, "113") = 1 Then
        cel.Offset(0, -1).Value = "Habshan"
        
        ElseIf InStr(1, cel.Value, "112") = 1 Then
        cel.Offset(0, -1).Value = "Habshan"
        
        ElseIf InStr(1, cel.Value, "81") = 1 Then
        cel.Offset(0, -1).Value = "Habshan-Maqta"
            
        ElseIf InStr(1, cel.Value, "77") = 1 Then
        cel.Offset(0, -1).Value = "Ruwais"
            
        ElseIf InStr(1, cel.Value, "51") = 1 Then
        cel.Offset(0, -1).Value = "Bab-Maqta"
            
        ElseIf InStr(1, cel.Value, "45") = 1 Then
        cel.Offset(0, -1).Value = "Habshan-Bab"
            
        ElseIf InStr(1, cel.Value, "33") = 1 Then
        cel.Offset(0, -1).Value = "Ruwais"
            
        ElseIf InStr(1, cel.Value, "26") = 1 Then
        cel.Offset(0, -1).Value = "Habshan"
            
        ElseIf InStr(1, cel.Value, "19") = 1 Then
        cel.Offset(0, -1).Value = "Habshan-Bab"
            
        ElseIf InStr(1, cel.Value, "18") = 1 Then
        cel.Offset(0, -1).Value = "Habshan-Ruwais"

        ElseIf InStr(1, cel.Value, "15") = 1 Then
        cel.Offset(0, -1).Value = "Habshan"
            
        ElseIf InStr(1, cel.Value, "13") = 1 Then
        cel.Offset(0, -1).Value = "Habshan"
            
        ElseIf InStr(1, cel.Value, "12") = 1 Then
        cel.Offset(0, -1).Value = "Habshan"
        
        ElseIf InStr(1, cel.Value, "0") = 1 Then
        cel.Offset(0, -1).Value = "'000"

        End If
   
    Next cel
    
End Sub
