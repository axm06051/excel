Option Explicit

Function RoundToList(value As Double) As Double
    Dim hours As Integer
    Dim minutes As Double
    Dim result As Double
    
    hours = Int(value)
    minutes = (value - hours) * 60

    Select Case minutes
        Case 1 To 6
            result = 0.1
        Case 7 To 9
            result = 0.15
        Case 10 To 12
            result = 0.2
        Case 13 To 15
            result = 0.25
        Case 16 To 30
            result = 0.5
        Case 31 To 45
            result = 0.75
        Case 46 To 60
            result = 1
    End Select

    RoundToList = hours + result
End Function
Sub Commit()
    Dim today As Date
    today = Date
    SaveTimesheetAndSummary today
End Sub
Sub SaveTimesheetAndSummary(Optional saveDate As Date)
    Dim ws As Worksheet
    Dim wsTimesheet As Worksheet
    Dim wsSummary As Worksheet
    Dim wbNew As Workbook
    Dim rng As Range
    Dim savePath As String
    Dim lastRow As Long
    Dim today As String
    Dim tbl As ListObject
    Dim tblRng As Range
    Dim summaryRng As Range
    
    If IsMissing(saveDate) Then
        saveDate = Date
    End If
    
    Set ws = ThisWorkbook.Sheets("Timesheet")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    Set rng = ws.Range("A1:G" & lastRow)
    
    GenerateSummary
    
    Set wsSummary = ThisWorkbook.Sheets("Summary")
    Set summaryRng = wsSummary.UsedRange
    
    Set wbNew = Workbooks.Add
    Set wsTimesheet = wbNew.Sheets(1)
    wsTimesheet.Name = "Timesheet"

    rng.Copy
    wsTimesheet.Range("A1").PasteSpecial Paste:=xlPasteValues
    
    wsTimesheet.Columns("A").NumberFormat = "m/d/yyyy h:mm"
    
    Set wsSummary = wbNew.Sheets.Add(After:=wsTimesheet)
    wsSummary.Name = "Summary"
    
    summaryRng.Copy
    wsSummary.Range("A1").PasteSpecial Paste:=xlPasteValues
    wsSummary.Range("A1").PasteSpecial Paste:=xlPasteFormats

    Set tblRng = wsTimesheet.Range("A1:G" & lastRow)
    Set tbl = wsTimesheet.ListObjects.Add(xlSrcRange, tblRng, , xlYes)
    tbl.TableStyle = "TableStyleMedium2"
    
    Set tblRng = wsSummary.UsedRange
    Set tbl = wsSummary.ListObjects.Add(xlSrcRange, tblRng, , xlYes)
    tbl.TableStyle = "TableStyleMedium2"

    wsTimesheet.Columns.AutoFit
    wsTimesheet.Rows.AutoFit
    wsSummary.Columns.AutoFit
    wsSummary.Rows.AutoFit

    today = Format(saveDate, "yyyy-mm-dd")
    savePath = Application.DefaultFilePath & "\Timesheets\" & today & ".xlsx"
    
    wbNew.SaveAs Filename:=savePath, fileFormat:=xlOpenXMLWorkbook
    wbNew.Close
    
    MsgBox "Timesheet and summary saved as " & savePath, vbInformation
End Sub
Sub ClearTimesheet()
    On Error GoTo ErrorHandler

    Dim today As Date
    Dim prevWorkingDay As Date
    Dim todayString As String
    Dim savePath As String

    today = Date
    todayString = Format(today, "yyyy-mm-dd")

    Select Case Weekday(today, vbMonday)
        Case 1 ' Monday
            prevWorkingDay = today - 3 ' Previous Friday
        Case 2 ' Tuesday
            prevWorkingDay = today - 1 ' Previous Monday
        Case 3 To 6 ' Wednesday to Friday
            prevWorkingDay = today - 1 ' Previous day
        Case 7 ' Saturday
            prevWorkingDay = today - 1 ' Previous Friday
        Case 8 ' Sunday
            prevWorkingDay = today - 2 ' Previous Friday
    End Select

    Dim prevWorkingDayString As String
    prevWorkingDayString = Format(prevWorkingDay, "yyyy-mm-dd")
    savePath = Application.DefaultFilePath & "\Timesheets\" & prevWorkingDayString & ".xlsx"

    If Dir(savePath) = "" Then
        SaveTimesheetAndSummary prevWorkingDay
        If Dir(savePath) = "" Then
            Err.Raise vbObjectError + 513, , "Previous working day's file not found and unable to create it."
        End If
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Timesheet")
    
    ws.Rows("2:" & ws.Rows.Count).ClearContents
    ws.ListObjects(1).Resize ws.Range("A1:G2")
    
    With ws
        .Range("E2").Formula = _
            "=IF(AND(ISBLANK(A2), ISBLANK(A3)), """", IF(ISBLANK(A2), """", TEXT(INT((IF(ISBLANK(A3), NOW(), A3) - A2) * 24), ""0"") & "" hrs "" & TEXT(INT((IF(ISBLANK(A3), NOW(), A3) - A2) * 1440) - INT((IF(ISBLANK(A3), NOW(), A3) - A2) * 24) * 60, ""0"") & "" minutes""))"

        .Range("F2").Formula = _
            "=IF(AND(ISBLANK(A2), ISBLANK(A3)), """", ROUND((IF(A3="""",NOW(),A3) - A2) * 24, 2))"
    End With

    Dim summarySheet As Worksheet
    On Error Resume Next
    Set summarySheet = ThisWorkbook.Sheets("Summary")
    On Error GoTo ErrorHandler

    If Not summarySheet Is Nothing Then
        Application.DisplayAlerts = False
        summarySheet.Delete
        Application.DisplayAlerts = True
    End If

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    Application.DisplayAlerts = True
End Sub
Sub GenerateSummary()
    Dim ws As Worksheet
    Dim wsSummary As Worksheet
    Dim lastRow As Long
    Dim summaryRow As Long
    Dim summaryRng As Range
    Dim tbl As ListObject
    Dim tblRng As Range
    Dim totalSum As Double
    Dim catSum As Double
    Dim eodSum As Double
    Dim lunchBreakSum As Double
    Dim cell As Range
    Dim categories As Collection
    Dim comments As Collection
    Dim cases As Collection
    Dim cat As Variant
    Dim com As Variant
    Dim cas As Variant
    Dim lunchBreakDetails As Object
    Dim categoryTotals As Object
    
    Set ws = ThisWorkbook.Sheets("Timesheet")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    Set summaryRng = ws.Range("A1:F" & lastRow)
    
    On Error Resume Next
    Set wsSummary = ThisWorkbook.Sheets("Summary")
    If wsSummary Is Nothing Then
        Set wsSummary = ThisWorkbook.Sheets.Add(After:=ws)
        wsSummary.Name = "Summary"
    Else
        wsSummary.Cells.Clear
    End If
    On Error GoTo 0
    
    For Each cell In ws.Range("C2:C" & lastRow)
        If Len(Trim(cell.value)) = 0 Then
            cell.value = "Lunch/Break"
        ElseIf InStr(1, cell.value, "EOD", vbTextCompare) > 0 Then
            cell.value = "EOD"
        End If
    Next cell
    
    summaryRow = 1
    wsSummary.Cells(summaryRow, 1).value = "Category"
    wsSummary.Cells(summaryRow, 2).value = "Total SalesForce Entry"
    wsSummary.Cells(summaryRow, 3).value = "Detail Type"
    wsSummary.Cells(summaryRow, 4).value = "Detail"
    wsSummary.Cells(summaryRow, 5).value = "Detail Total"
    summaryRow = summaryRow + 1
    
    totalSum = 0
    lunchBreakSum = 0
    eodSum = 0
    Set lunchBreakDetails = CreateObject("Scripting.Dictionary")
    Set categoryTotals = CreateObject("Scripting.Dictionary")
    
    For Each cell In ws.Range("F2:F" & lastRow)
        If Len(cell.Offset(0, -2).value) > 0 Then
            Dim roundedValue As Double
            roundedValue = RoundToList(cell.value)
            
            If cell.Offset(0, -3).value <> "Lunch/Break" And _
               cell.Offset(0, -3).value <> "EOD" Then
                totalSum = totalSum + roundedValue
                If Not categoryTotals.Exists(cell.Offset(0, -3).value) Then
                    categoryTotals.Add cell.Offset(0, -3).value, roundedValue
                Else
                    categoryTotals(cell.Offset(0, -3).value) = categoryTotals(cell.Offset(0, -3).value) + roundedValue
                End If
            ElseIf cell.Offset(0, -3).value = "Lunch/Break" Then
                lunchBreakSum = lunchBreakSum + roundedValue
                Dim comment As String
                comment = Trim(cell.Offset(0, -2).value)
                If Not lunchBreakDetails.Exists(comment) Then
                    lunchBreakDetails.Add comment, roundedValue
                Else
                    lunchBreakDetails(comment) = lunchBreakDetails(comment) + roundedValue
                End If
            ElseIf cell.Offset(0, -3).value = "EOD" Then
                eodSum = eodSum + roundedValue
            End If
        End If
    Next cell
    
    ' Subtract Lunch/Break from totalSum
    totalSum = totalSum - lunchBreakSum
    ' Adjust lunchBreakSum to appear as negative in the summary
    lunchBreakSum = -lunchBreakSum
    
    wsSummary.Cells(summaryRow, 1).value = "Overall Total (excluding Lunch/Break)"
    wsSummary.Cells(summaryRow, 2).value = totalSum
    summaryRow = summaryRow + 1
    
    wsSummary.Cells(summaryRow, 1).value = "Sum by Category"
    summaryRow = summaryRow + 1
    
    Set categories = New Collection
    On Error Resume Next
    For Each cell In summaryRng.Columns(3).Cells
        If cell.row > 1 Then
            categories.Add cell.value, CStr(cell.value)
        End If
    Next cell
    On Error GoTo 0
    
    For Each cat In categories
        If cat <> "Lunch/Break" And cat <> "EOD" Then
            wsSummary.Cells(summaryRow, 1).value = cat
            wsSummary.Cells(summaryRow, 2).value = categoryTotals(cat)
            summaryRow = summaryRow + 1
            
            If cat = "Support Work" Then
                wsSummary.Cells(summaryRow, 3).value = "Breakdown by Case Number"
                summaryRow = summaryRow + 1
                Set cases = New Collection
                On Error Resume Next
                For Each cell In summaryRng.Columns(2).Cells
                    If cell.row > 1 And cell.Offset(0, 1).value = cat Then
                        cases.Add cell.value, CStr(cell.value)
                    End If
                Next cell
                On Error GoTo 0
                For Each cas In cases
                    wsSummary.Cells(summaryRow, 4).value = cas
                    wsSummary.Cells(summaryRow, 5).value = RoundToList(WorksheetFunction.SumIfs(summaryRng.Columns(6), summaryRng.Columns(3), cat, summaryRng.Columns(2), cas))
                    summaryRow = summaryRow + 1
                Next cas
            ElseIf cat = "AMPP Support" Then
                ' Handling "AMPP Support" category
                Set comments = New Collection
                Set cases = New Collection
                
                On Error Resume Next
                For Each cell In summaryRng.Columns(4).Cells
                    If cell.row > 1 And cell.Offset(0, -1).value = cat Then
                        If Len(cell.Offset(0, -2).value) = 0 Then
                            comments.Add cell.value, CStr(cell.value)
                        Else
                            cases.Add cell.Offset(0, -2).value, CStr(cell.Offset(0, -2).value)
                        End If
                    End If
                Next cell
                On Error GoTo 0
                
                ' Process comments
                If comments.Count > 0 Then
                    wsSummary.Cells(summaryRow, 3).value = "Breakdown by Comment if Case Number is empty"
                    summaryRow = summaryRow + 1
                    For Each com In comments
                        wsSummary.Cells(summaryRow, 4).value = com
                        wsSummary.Cells(summaryRow, 5).value = RoundToList(WorksheetFunction.SumIfs(summaryRng.Columns(6), summaryRng.Columns(3), cat, summaryRng.Columns(4), com, summaryRng.Columns(2), ""))
                        summaryRow = summaryRow + 1
                    Next com
                End If
                
                ' Process cases
                If cases.Count > 0 Then
                    wsSummary.Cells(summaryRow, 3).value = "Breakdown by Case Number"
                    summaryRow = summaryRow + 1
                    For Each cas In cases
                        wsSummary.Cells(summaryRow, 4).value = cas
                        wsSummary.Cells(summaryRow, 5).value = RoundToList(WorksheetFunction.SumIfs(summaryRng.Columns(6), summaryRng.Columns(3), cat, summaryRng.Columns(2), cas))
                        summaryRow = summaryRow + 1
                    Next cas
                End If
                
            ElseIf cat = "Internal Admin" Or cat = "Customer Admin" Or cat = "Personal Development" Then
                ' Handling other categories
                Set comments = New Collection
                
                On Error Resume Next
                For Each cell In summaryRng.Columns(4).Cells
                    If cell.row > 1 And cell.Offset(0, -1).value = cat Then
                        comments.Add cell.value, CStr(cell.value)
                    End If
                Next cell
                On Error GoTo 0
                
                wsSummary.Cells(summaryRow, 3).value = "Breakdown by Comment"
                summaryRow = summaryRow + 1
                For Each com In comments
                    wsSummary.Cells(summaryRow, 4).value = com
                    wsSummary.Cells(summaryRow, 5).value = RoundToList(WorksheetFunction.SumIfs(summaryRng.Columns(6), summaryRng.Columns(3), cat, summaryRng.Columns(4), com))
                    summaryRow = summaryRow + 1
                Next com
            End If
        End If
    Next cat
    
    ' Lunch/Break section
    wsSummary.Cells(summaryRow, 1).value = "Lunch/Break"
    wsSummary.Cells(summaryRow, 2).value = lunchBreakSum
    summaryRow = summaryRow + 1
    
    ' Add breakdown for Lunch/Break
    wsSummary.Cells(summaryRow, 3).value = "Breakdown of Lunch/Break"
    summaryRow = summaryRow + 1
    
    Dim key As Variant
    For Each key In lunchBreakDetails.Keys
        wsSummary.Cells(summaryRow, 4).value = key
        wsSummary.Cells(summaryRow, 5).value = lunchBreakDetails(key)
        summaryRow = summaryRow + 1
    Next key
    
    ' EOD section
    wsSummary.Cells(summaryRow, 1).value = "EOD"
    wsSummary.Cells(summaryRow, 2).value = eodSum
    summaryRow = summaryRow + 1
    
    Set tblRng = wsSummary.Range("A1:E" & summaryRow - 1)
    Set tbl = wsSummary.ListObjects.Add(xlSrcRange, tblRng, , xlYes)
    tbl.TableStyle = "TableStyleMedium2"
    
    wsSummary.Columns("B:B").NumberFormat = "#,##0.00"
    wsSummary.Columns("E:E").NumberFormat = "#,##0.00"
    
    ' Autofit columns
    wsSummary.Columns("A:E").AutoFit
    
    wsSummary.Activate
    wsSummary.Cells(1, 1).Select
End Sub
