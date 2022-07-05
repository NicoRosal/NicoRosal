Option Explicit

Sub Button1_Click()

    Dim myDataRng As Range
    Set myDataRng = Range("A1:A" & Cells(Rows.Count, "A").End(xlUp).Row)
    
    Dim cell As Range
    
    For Each cell In myDataRng
        GetFileName cell.Row, cell(1, 1)
    Next cell
    
    Set myDataRng = Nothing
    
End Sub

Private Sub GetFileName(iRow As Integer, sFilePath As String)
    On Error GoTo ErrHandler
    
    ' EXTRACT THE FILENAME FROM A FILE PATH.
    Dim objFSO
    Set objFSO = CreateObject("scripting.filesystemobject")
    Dim fileName As String
    fileName = objFSO.GetFileName(sFilePath)
    
    Cells(iRow, 2).Value = fileName         ' THE SECOND COLUMN.
     
ErrHandler:
    '
End Sub
