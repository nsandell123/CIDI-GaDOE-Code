Global totalRegistered As Variant
Global totalRequested As Variant
Global totalCategories() As Variant
Sub findNewDistricts()
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
    
    Dim categories() As Variant
    'Local Districts, State Schools, Charters, GNETS
    ReDim categories(4)
    
    Dim cellplaceholder As Range
    Dim finalString As String
    finalString = finalString + "NEW DISTRICTS REGISTERED: " + vbNewLine
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
    If StrComp(finalString, "NEW DISTRICTS REGISTERED: " + vbNewLine, vbTextCompare) = 0 Then
        MsgBox "No new districts have registered in the portal"
    Else
        MsgBox finalString
    End If
    Dim LRowRequests1 As Long
    Dim LRowRequests2 As Long
    LRowRequests1 = Worksheets(sheetNames(2)).Cells(Rows.Count, "M").End(xlUp).Row
    LRowRequests2 = Worksheets(sheetNames(3)).Cells(Rows.Count, "M").End(xlUp).Row
    
    Dim RequestsRange1 As Range
    Dim RequestsRange2 As Range
    
    Set RequestsRange1 = Worksheets(sheetNames(2)).Range("M1:" + "M" + Right(Str(LRowRequests1), Len(LRowRequests1) - 1))
    Set RequestsRange2 = Worksheets(sheetNames(3)).Range("M1:" + "M" + Right(Str(LRowRequests2), Len(LRowRequests2) - 1))
    
    
    Dim finalStringRequests As String
    finalStringRequests = finalStringRequests + "NEW DISTRICTS REQUESTING: " + vbNewLine
    For Each cellplaceholder In RequestsRange1
        If WorksheetFunction.CountIf(RequestsRange2, cellplaceholder) = 0 And cellplaceholder.Value <> 0 Then
            requestedPortal = requestedPortal + 1
            finalStringRequests = finalStringRequests + cellplaceholder.Value + cellplaceholder.Offset(0, 1).Value + vbNewLine
        End If
    Next cellplaceholder
    If StrComp(finalStringRequests, "NEW DISTRICTS REQUESTING: " + vbNewLine, vbTextCompare) = 0 Then
        MsgBox "No new districts in the portal that are requesting software"
    Else
        MsgBox finalStringRequests
    End If
    
    totalRegistered = registeredPortal
    totalRequested = requestedPortal
    ReDim Preserve totalCategories(UBound(categories) - LBound(categories) + 1)
    totalCategories = categories

End Sub
