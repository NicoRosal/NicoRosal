Sub consolidateValues()

    Dim sh As Worksheet
    Dim rw As Range
    Dim s As String
    Dim i As Integer

    Set sh = ThisWorkbook.Sheets("Sheet1")

    For Each rw In Intersect(sh.UsedRange, sh.Range("A:B")).Rows

        'Skip row 1 (assumed headers)
        If rw.Row <> 1 Then

            s = ""

            For i = sh.UsedRange.Rows.Count To rw.Row + 1 Step -1

                If rw.Cells(1, 1) = sh.Cells(i, 1) Then
                    s = sh.Cells(i, 2).Value & IIf(s = "", "", ",") & s
                    sh.Rows(i).Delete
                End If

            Next i

            If s <> "" Then rw.Cells(1, 2).Value = rw.Cells(1, 2).Value & "," & s

        End If

    Next rw

End Sub
