Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    ' Call the CountRedCells macro for each sheet before saving
    Dim ws As Worksheet
    Dim rowCount As Integer
    rowCount = 1
    
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "ResultSheet" Then
            CountRedCells ws, rowCount
            rowCount = rowCount + 1 ' Move to the next row for the next sheet
        End If
    Next ws
End Sub

Sub CountRedCells(ws As Worksheet, rowCount As Integer)
    Dim sourceRange As Range
    Dim cell As Range
    Dim redCount As Long
    
    ' Set the source range for the specific sheet
    Set sourceRange = ws.UsedRange ' You can adjust the range as needed
    
    ' Create or reference the result worksheet
    On Error Resume Next
    Dim wsResult As Worksheet
    Set wsResult = ThisWorkbook.Sheets("ResultSheet")
    On Error GoTo 0
    
    If wsResult Is Nothing Then
        ' Create a new worksheet if it doesn't exist
        Set wsResult = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsResult.Name = "ResultSheet"
    End If
    
    ' Reset redCount
    redCount = 0
    
    ' Loop through each cell in the source range
    For Each cell In sourceRange
        If cell.Interior.Color = RGB(255, 0, 0) Then ' Check for red fill color
            redCount = redCount + 1
        End If
    Next cell
    
    ' Display the count in the result worksheet
    wsResult.Cells(rowCount, 1).Value = ws.Name
    wsResult.Cells(rowCount, 2).Value = redCount
    
    ' Optionally, you can format the result sheet or cells as needed
End Sub

