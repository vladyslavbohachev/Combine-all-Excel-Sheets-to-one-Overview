Sub CombineSheets()
    Dim ws As Worksheet
    Dim wsOverview As Worksheet
    Dim lRow As Long
    Set wsOverview = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsOverview.Name = "Overview"

    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "Overview" Then
            lRow = wsOverview.Cells(wsOverview.Rows.Count, "A").End(xlUp).Row
            'here you can specify which range/s you need....
            ws.Range("B8").Copy wsOverview.Cells(lRow + 1, "A")
            ws.Range("B14").Copy wsOverview.Cells(lRow + 1, "B")
            ws.Range("B15").Copy wsOverview.Cells(lRow + 1, "C")
            ws.Range("B16").Copy wsOverview.Cells(lRow + 1, "D")
        End If
    Next ws
End Sub

