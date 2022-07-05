Sub Macro1()
Dim mbResult As Integer
mbResult = MsgBox("These changes cannot be undone. Would you like to save a copy before proceeding?", _
 vbYesNoCancel)

Select Case mbResult
    Case vbYes
    With ActiveWorkbook
    Rows("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
        End With
    Case vbNo
     'No
    Case vbCancel
    
        Exit Sub

End Select


End Sub

