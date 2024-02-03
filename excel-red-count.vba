Sub CountRedCells()
    Dim wsSource As Worksheet
    Dim wsResult As Worksheet
    Dim sourceRange As Range
    Dim cell As Range
    Dim redCount As Long
    
    ' Set the source worksheet and range
    Set wsSource = ThisWorkbook.Sheets("Sheet1") ' Change "Sheet1" to the name of your source sheet
    Set sourceRange = wsSource.Range("A1:F10") ' Change the range as needed
    
    ' Create or reference the result worksheet
    On Error Resume Next
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
    wsResult.Range("A1").Value = wsSource.Name
    wsResult.Range("B1").Value = redCount
    
    ' Optionally, you can format the result sheet or cells as needed
    
    ' Activate the source sheet (optional)
    wsSource.Activate
End Sub

