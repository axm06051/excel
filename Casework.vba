Sub ArchiveRow()
    Dim wsCases As Worksheet
    Dim wsArchive As Worksheet
    Dim tblCaseTracker As ListObject
    Dim tblCasesArchive As ListObject
    Dim selectedRow As ListRow
    Dim newRow As ListRow
    Dim i As Integer
    
    Set wsCases = ThisWorkbook.Sheets("Cases")
    Set wsArchive = ThisWorkbook.Sheets("Archive")
    
    Set tblCaseTracker = wsCases.ListObjects("CaseTracker")
    Set tblCasesArchive = wsArchive.ListObjects("CasesArchive")
    
    Set selectedRow = tblCaseTracker.ListRows(Application.ActiveCell.row - tblCaseTracker.HeaderRowRange.row)
    
    Set newRow = tblCasesArchive.ListRows.Add
    For i = 1 To tblCaseTracker.ListColumns.Count
        newRow.Range(1, i).value = selectedRow.Range(1, i).value
    Next i
    
    selectedRow.Delete
End Sub
Sub UnarchiveRow()
    Dim wsCases As Worksheet
    Dim wsArchive As Worksheet
    Dim tblCaseTracker As ListObject
    Dim tblCasesArchive As ListObject
    Dim selectedRow As ListRow
    Dim newRow As ListRow
    Dim i As Integer
    
    Set wsCases = ThisWorkbook.Sheets("Cases")
    Set wsArchive = ThisWorkbook.Sheets("Archive")
    
    Set tblCaseTracker = wsCases.ListObjects("CaseTracker")
    Set tblCasesArchive = wsArchive.ListObjects("CasesArchive")
    
    Set selectedRow = tblCasesArchive.ListRows(Application.ActiveCell.row - tblCasesArchive.HeaderRowRange.row)
    
    Set newRow = tblCaseTracker.ListRows.Add
    For i = 1 To tblCasesArchive.ListColumns.Count
        newRow.Range(1, i).value = selectedRow.Range(1, i).value
    Next i
    
    selectedRow.Delete
End Sub
Sub AdjustRowHeights()
    Dim tbl As ListObject
    Dim rowHeight As Double
    Dim padding As Double
    Dim maxRowHeight As Double
    Dim targetRow As Range
    Dim targetCell As Range
    Dim rowHeights() As Double
    Dim i As Long

    On Error Resume Next
    Set tbl = ThisWorkbook.Sheets("Cases").ListObjects("CaseTracker")
    On Error GoTo 0

    If tbl Is Nothing Then Exit Sub

    padding = 0.5 * Application.StandardFontSize

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ReDim rowHeights(1 To tbl.DataBodyRange.Rows.Count)

    i = 1
    For Each targetRow In tbl.DataBodyRange.Rows
        maxRowHeight = 0
        For Each targetCell In targetRow.Cells
            targetCell.WrapText = True
            targetCell.EntireRow.AutoFit
            maxRowHeight = Application.Max(maxRowHeight, targetCell.Height)
        Next targetCell

        rowHeights(i) = maxRowHeight + 2 * padding
        i = i + 1
    Next targetRow

    i = 1
    For Each targetRow In tbl.DataBodyRange.Rows
        targetRow.rowHeight = rowHeights(i)
        i = i + 1
    Next targetRow

    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

