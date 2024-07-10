Private Sub Worksheet_Change(ByVal Target As Range)
    Dim tbl As ListObject
    Dim updatedCol As ListColumn
    Dim changedCell As Range
    Dim updateColIndex As Long

    On Error Resume Next
    Set tbl = Me.ListObjects("CaseTracker")
    On Error GoTo 0
    
    If tbl Is Nothing Then Exit Sub

    updateColIndex = tbl.ListColumns("updated_on").Index

    If Not Intersect(Target, tbl.DataBodyRange) Is Nothing Then
        For Each changedCell In Target
            If changedCell.Column <> updateColIndex Then
                tbl.DataBodyRange.Cells(changedCell.row - tbl.DataBodyRange.row + 1, updateColIndex).value = Format(Now, "yyyy-mm-dd hh:mm")
            End If
        Next changedCell
    End If
    Call AdjustRowHeights
End Sub
