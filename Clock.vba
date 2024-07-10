Dim NextTick As Date
Sub StartTimer()
    ' Schedule the next update at the beginning of the next minute
    NextTick = Now + TimeValue("00:00:01") - TimeSerial(0, 0, Second(Now)) + TimeValue("00:01:00")
    Application.OnTime NextTick, "Recalc", , True
End Sub
Sub Recalc()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Timesheet")
    
    ws.ListObjects("Timesheet").Range.Calculate
    
    StartTimer
End Sub
Sub StopTimer()
    On Error Resume Next
    Application.OnTime NextTick, "Recalc", , False
End Sub
Private Sub Workbook_Open()
    StartTimer
End Sub
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    StopTimer
End Sub
Sub CaptureStartTime()
    Dim currentDateTime As String
    currentDateTime = Format(Now, "yyyy-mm-dd hh:mm")
    ActiveCell.value = currentDateTime
End Sub


