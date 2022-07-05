'vlookup to another workbook ex
Sub AddData()

On Error Resume Next ' I also suggest removing this since it wont warn you on an error.

Dim wb as Workbook
Dim wbExternal as Workbook

Dim ws as Worksheet
Dim wsExternal as Worksheet

'Open External Data Source
Set wbExternal = Workbooks.Open Filename:= _
    "W:\USB\Reporting\Book Tool\Attachments\Team Data.xls"

' Depending on the location of your file, you may run into issues with workbook.Open
' If this does become an issue, I tend to use Workbook.FollowHyperlink()


'View sheet where data will go into
' Windows("Book Tool - Updated Feb. 2017.xlsb").Activate
' Set wb = ActiveWorkbook

' As noted by Shai Rado, do this instead:
Se wb = Workbooks("Book Tool - Updated Feb. 2017.xlsb")

' Or if the workbook running the code is book tool
' Set wb = ThisWorkbook

 'Gets last row of Tool sheet
 Set ws = wb.Sheets("Book")
 lastrow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

'Lookup in External File
Set wsExternal = wbExternal.Sheets("Book")
wsExternal.Range("DE2:DE" & lastrow).FormulaR1C1 = "=VLOOKUP(RC[-108],'[Team Data.xls]SICcode'!C[-109]:C[-104],5,FALSE)"


'Close External Data File

ThisWorkbook.Saved = True
Application.DisplayAlerts = False
Windows("Team Data.xls").Close


MsgBox "Data Add Done"


End Sub