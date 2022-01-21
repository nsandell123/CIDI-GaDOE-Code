Sub UpdateTable()
    Worksheets(sheetNames(1)).Select
    'TxtHelp Rept
    Dim LRowFirst As Long
    LRowFirst = Worksheets(sheetNames(1)).Cells(Rows.Count, "K").End(xlUp).Row
    
    Dim oldDate As String
    oldDate = Cells(LRowFirst, 11).Offset(-6, 0)
    Worksheets(sheetNames(1)).Cells(LRowFirst, 11).Offset(-5, 0).Value = oldDate
    Worksheets(sheetNames(1)).Cells(LRowFirst, 11).Offset(-10, 0).Value = oldDate
    
    
    
    
    'Change the Dates
    Dim fetchedDate As String
    '123121
    fetchedDate = Right(Worksheets(sheetNames(1)).Name, 6)
    '1231
    fetchedDate = Left(fetchedDate, 4)
    fetchedDate = Left(fetchedDate, 2) & "/" & Right(fetchedDate, 2) & " Totals"
    
    Worksheets(sheetNames(1)).Cells(LRowFirst, 11).Offset(-6, 0).Value = fetchedDate
    Worksheets(sheetNames(1)).Cells(LRowFirst, 11).Offset(-6, 0).NumberFormat = "m/d"
    Worksheets(sheetNames(1)).Cells(LRowFirst, 11).Offset(-11, 0).Value = fetchedDate
    Worksheets(sheetNames(1)).Cells(LRowFirst, 11).Offset(-11, 0).NumberFormat = "m/d"
    
    
    'Actually Copy the Values
    Worksheets(sheetNames(1)).Cells(LRowFirst, 11).Offset(-5, 1) = Worksheets(sheetNames(1)).Cells(LRowFirst, 11).Offset(-6, 1)
    Worksheets(sheetNames(1)).Cells(LRowFirst, 11).Offset(-5, 2) = Worksheets(sheetNames(1)).Cells(LRowFirst, 11).Offset(-6, 2)
    Worksheets(sheetNames(1)).Cells(LRowFirst, 11).Offset(-5, 3) = Worksheets(sheetNames(1)).Cells(LRowFirst, 11).Offset(-6, 3)
    Worksheets(sheetNames(1)).Cells(LRowFirst, 11).Offset(-5, 4) = Worksheets(sheetNames(1)).Cells(LRowFirst, 11).Offset(-6, 4)
    
    
    Worksheets(sheetNames(1)).Cells(LRowFirst, 11).Offset(-10, 1) = Worksheets(sheetNames(1)).Cells(LRowFirst, 11).Offset(-6, 1)
    Worksheets(sheetNames(1)).Cells(LRowFirst, 11).Offset(-10, 2) = Worksheets(sheetNames(1)).Cells(LRowFirst, 11).Offset(-6, 2)
    Worksheets(sheetNames(1)).Cells(LRowFirst, 11).Offset(-10, 3) = Worksheets(sheetNames(1)).Cells(LRowFirst, 11).Offset(-6, 3)
    Worksheets(sheetNames(1)).Cells(LRowFirst, 11).Offset(-10, 4) = Worksheets(sheetNames(1)).Cells(LRowFirst, 11).Offset(-6, 4)
    
    Worksheets(sheetNames(1)).Cells(LRowFirst, 11).Offset(-6, 1) = Worksheets(sheetNames(1)).Cells(LRowFirst, 11).Offset(-11, 1)
    Worksheets(sheetNames(1)).Cells(LRowFirst, 11).Offset(-6, 2) = Worksheets(sheetNames(1)).Cells(LRowFirst, 11).Offset(-11, 2)
    Worksheets(sheetNames(1)).Cells(LRowFirst, 11).Offset(-6, 3) = Worksheets(sheetNames(1)).Cells(LRowFirst, 11).Offset(-11, 3)
    Worksheets(sheetNames(1)).Cells(LRowFirst, 11).Offset(-6, 4) = Worksheets(sheetNames(1)).Cells(LRowFirst, 11).Offset(-11, 4)
    
    
    
    
    
    
    
    
    
End Sub
