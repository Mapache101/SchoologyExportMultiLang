Sub UpdateCategoryAveragesFormatting()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim avgCol As Range
    Dim cell As Range
    
    ' Set worksheet
    Set ws = ActiveSheet
    
    ' Find last row with data
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Find last column with data
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Loop through each column where category averages are calculated
    For Each avgCol In ws.Range(ws.Cells(1, 2), ws.Cells(1, lastCol))
        ' Check if the column contains an average header
        If InStr(1, avgCol.Value, "Average", vbTextCompare) > 0 Then
            ' Apply light blue color to the column header
            avgCol.Interior.Color = RGB(173, 216, 230)
            
            ' Apply formatting to populated cells
            For Each cell In ws.Range(ws.Cells(2, avgCol.Column), ws.Cells(lastRow, avgCol.Column))
                If IsNumeric(cell.Value) And Not IsEmpty(cell.Value) Then
                    cell.Interior.Color = RGB(173, 216, 230) ' Light blue background
                Else
                    cell.Interior.ColorIndex = xlNone ' Keep empty cells white
                End If
            Next cell
        End If
    Next avgCol
    
End Sub
