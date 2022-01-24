Sub UpdateRegisterSheets()
    Dim registeredPortal As Long
    registeredPortal = 0
    Dim requestedPortal As Long
    requestedPortal = 0
    
    Dim LRowUsers1 As Long
    Dim LRowUsers2 As Long
    LRowUsers1 = Worksheets(sheetNames(4)).Cells(Rows.Count, "G").End(xlUp).Row
    LRowUsers2 = Worksheets(sheetNames(5)).Cells(Rows.Count, "G").End(xlUp).Row
    
    Dim UserRange1 As Range
    Dim UserRange2 As Range
    
    Set UserRange1 = Worksheets(sheetNames(4)).Range("G1:" + "G" + Right(Str(LRowUsers1), Len(LRowUsers1) - 1))
    Set UserRange2 = Worksheets(sheetNames(5)).Range("G1:" + "G" + Right(Str(LRowUsers2), Len(LRowUsers2) - 1))
    
    Dim categories() As Long
    'Local Districts, State Schools, Charters, GNETS
    ReDim categories(4)
    
    Dim cellplaceholder As Range
    Dim finalString As String
    For Each cellplaceholder In UserRange1
        If WorksheetFunction.CountIf(UserRange2, cellplaceholder) = 0 And cellplaceholder.Value <> 0 Then
            registeredPortal = registeredPortal + 1
            finalString = finalString + cellplaceholder.Value + cellplaceholder.Offset(0, -1).Value + vbNewLine
            If IsNumeric(Left(cellplaceholder.Offset(0, -1).Value, 1)) Then
                categories(0) = categories(0) + 1
            End If
            If StrComp(cellplaceholder.Offset(0, -1).Value, "State Schools") = 0 Then
                categories(1) = categories(1) + 1
            End If
            If StrComp(cellplaceholder.Offset(0, -1).Value, "Charter Schools") = 0 Then
                categories(2) = categories(2) + 1
            End If
            If StrComp(cellplaceholder.Offset(0, -1).Value, "GNETS") = 0 Then
                categories(3) = categories(3) + 1
            End If
        End If
    Next cellplaceholder
    
    Dim LRowRequests1 As Long
    Dim LRowRequests2 As Long
    LRowRequests1 = Worksheets(sheetNames(2)).Cells(Rows.Count, "M").End(xlUp).Row
    LRowRequests2 = Worksheets(sheetNames(3)).Cells(Rows.Count, "M").End(xlUp).Row
    
    Dim RequestsRange1 As Range
    Dim RequestsRange2 As Range
    
    Set RequestsRange1 = Worksheets(sheetNames(2)).Range("M1:" + "M" + Right(Str(LRowRequests1), Len(LRowRequests1) - 1))
    Set RequestsRange2 = Worksheets(sheetNames(3)).Range("M1:" + "M" + Right(Str(LRowRequests2), Len(LRowRequests2) - 1))
    
    
    Dim finalStringRequests As String
    For Each cellplaceholder In RequestsRange1
        If WorksheetFunction.CountIf(RequestsRange2, cellplaceholder) = 0 And cellplaceholder.Value <> 0 Then
            requestedPortal = requestedPortal + 1
            finalStringRequests = finalStringRequests + cellplaceholder.Value + cellplaceholder.Offset(0, 1).Value + vbNewLine
        End If
    Next cellplaceholder
    
    'In Register Monthly, you have to update the Next with the totals from Sheet 1
    Dim nextPosition As Long
    nextPosition = 7
    While StrComp(Worksheets(sheetNames(6)).Cells(nextPosition, "A").Value, "Next") <> 0
        nextPosition = nextPosition + 1
    Wend
    
    Worksheets(sheetNames(6)).Cells(nextPosition, "B").Value = Worksheets(sheetNames(1)).Cells(294, "L").Value
    Worksheets(sheetNames(6)).Cells(nextPosition, "C").Value = Worksheets(sheetNames(1)).Cells(294, "M").Value
    Worksheets(sheetNames(6)).Cells(nextPosition, "D").Value = Worksheets(sheetNames(1)).Cells(294, "N").Value
    Worksheets(sheetNames(6)).Cells(nextPosition, "E").Value = Worksheets(sheetNames(1)).Cells(294, "O").Value
    
    nextPosition = nextPosition + 1
    While StrComp(Worksheets(sheetNames(6)).Cells(nextPosition, "A").Value, "Next") <> 0
        nextPosition = nextPosition + 1
    Wend
    
    Worksheets(sheetNames(6)).Cells(nextPosition, "B").Value = Worksheets(sheetNames(6)).Cells(nextPosition, "B").Value + registeredPortal
    Worksheets(sheetNames(6)).Cells(nextPosition, "G").Value = Worksheets(sheetNames(6)).Cells(nextPosition, "G").Value + requestedPortal
    
    Dim j As Long
    
    For j = 0 To 3
        Worksheets(sheetNames(6)).Cells(57 + j, "C").Value = Worksheets(sheetNames(6)).Cells(57 + j, "C").Value + categories(j)
    Next j
    
    
    
    Set FindRow = Worksheets(sheetNames(7)).Range("B:B").Find(What:="Requested", LookIn:=xlValues)
    Rows(FindRow.Row).Insert
    
    'Get the Date
    myValue = "01/24/2022"
    Dim standardDate As Variant
    standardDate = Left(CStr(myValue), 5)
    standardDate = standardDate + "Totals"
    Worksheets(sheetNames(7)).Cells(FindRow.Row - 1, "A").Value = standardDate
    Worksheets(sheetNames(7)).Cells(FindRow.Row - 1, "B").Value = Worksheets(sheetNames(1)).Cells(294, "L").Value
    Worksheets(sheetNames(7)).Cells(FindRow.Row - 1, "C").Value = Worksheets(sheetNames(1)).Cells(294, "M").Value
    Worksheets(sheetNames(7)).Cells(FindRow.Row - 1, "D").Value = Worksheets(sheetNames(1)).Cells(294, "N").Value
    Worksheets(sheetNames(7)).Cells(FindRow.Row - 1, "E").Value = Worksheets(sheetNames(1)).Cells(294, "O").Value
    
    
    

End Sub
