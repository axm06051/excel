Sub BoldTextInBrackets()
    Dim ws As Worksheet
    Dim cell As Range
    Dim startPos As Long
    Dim endPos As Long
    Dim text As String
    
    ' Loop through each worksheet in the active workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Loop through each cell with content in the worksheet
        For Each cell In ws.UsedRange
            ' Check if the cell contains text
            If Not IsEmpty(cell.value) And cell.HasFormula = False Then
                text = cell.value
                startPos = InStr(text, "[")
                endPos = InStr(text, "]")
                
                ' If both opening and closing brackets are found
                If startPos > 0 And endPos > startPos Then
                    ' Format the text between the brackets as bold
                    cell.Characters(startPos, endPos - startPos + 1).Font.Bold = True
                End If
            End If
        Next cell
    Next ws
End Sub

