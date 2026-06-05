'MACRO 1: Auto-Format All Sheets 
Sub FormatAllSheets()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Cells.Font.Name = "Arial"
        ws.Cells.Font.Size = 9
        ws.Rows(1).Font.Size = 12
        ws.Rows(1).Font.Bold = True
        ws.Cells.EntireColumn.AutoFit
    Next ws
    MsgBox "All sheets formatted successfully.", vbInformation
End Sub

'MACRO 2: Highlight Overpayments Red
Sub HighlightOverpayments()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim threshold As Double

    threshold = 4500
    ws = ThisWorkbook.Sheets("Overpayments")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = 4 To lastRow
        If ws.Cells(i, 5).Value > threshold Then
            ws.Rows(i).Interior.Color = RGB(255, 199, 199) 
            ws.Cells(i, 5).Font.Bold = True
            ws.Cells(i, 5).Font.Color = RGB(192, 0, 0)
        End If
    Next i

    MsgBox "Overpayments highlighted.", vbInformation
End Sub

'MACRO 3: Add Autofilter to Flagged Transactions
Sub AddFlagFilter()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Flagged Transactions")
    ws.Rows(3).AutoFilter
    MsgBox "Filter applied to Flagged Transactions sheet. Use dropdowns to filter by region, flag reason, etc.", vbInformation
End Sub

'MACRO 4: Export Flagged Transactions to New Workbook
Sub ExportFlaggedToNewWorkbook()
    Dim srcWs As Worksheet
    Dim newWb As Workbook
    Dim newWs As Worksheet

    Set srcWs = ThisWorkbook.Sheets("Flagged Transactions")
    Set newWb = Workbooks.Add
    Set newWs = newWb.Sheets(1)
    newWs.Name = "Flagged Review"

    srcWs.UsedRange.Copy
    newWs.Paste
    newWs.Cells.EntireColumn.AutoFit

    Dim savePath As String
    savePath = Environ("USERPROFILE") & "\Desktop\Flagged_Transactions_Review.xlsx"
    newWb.SaveAs savePath
    newWb.Close

    MsgBox "Exported flagged transactions to: " & savePath, vbInformation
End Sub

'MACRO 5: Generate Summary Statistics
Sub GenerateSummaryStats()
    Dim wb As Workbook
    Dim ws As Worksheet
    Set wb = ThisWorkbook

    'Delete if already exists
    On Error Resume Next
    Application.DisplayAlerts = False
    wb.Sheets("Quick Stats").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    ws.Name = "Quick Stats"

    'Header
    ws.Range("A1").Value = "Quick Statistics — Auto Generated"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 13
    ws.Range("A1").Font.Color = RGB(31, 78, 121)

    'Pull stats from Executive Summary 
    Dim sumWs As Worksheet
    Set sumWs = wb.Sheets("Executive Summary")

    ws.Range("A3").Value = "Metric"
    ws.Range("B3").Value = "Value"
    ws.Range("A3:B3").Font.Bold = True
    ws.Range("A3:B3").Interior.Color = RGB(31, 78, 121)
    ws.Range("A3:B3").Font.Color = RGB(255, 255, 255)

    Dim stats(1 To 5, 1 To 2) As String
    stats(1, 1) = "Report Generated"
    stats(1, 2) = Format(Now(), "MMMM DD, YYYY HH:MM")
    stats(2, 1) = "Total Transactions"
    stats(2, 2) = "=COUNTA('Flagged Transactions'!A:A)-3"
    stats(3, 1) = "Flagged Transactions"
    stats(3, 2) = "=COUNTA('Flagged Transactions'!A:A)-3"
    stats(4, 1) = "Sheets in Report"
    stats(4, 2) = CStr(wb.Sheets.Count - 1)
    stats(5, 1) = "Analyst"
    stats(5, 2) = Environ("USERNAME")

    Dim r As Integer
    For r = 1 To 5
        ws.Cells(r + 3, 1).Value = stats(r, 1)
        ws.Cells(r + 3, 2).Value = stats(r, 2)
    Next r

    ws.Columns("A:B").AutoFit
    MsgBox "Quick Stats sheet created.", vbInformation
End Sub
