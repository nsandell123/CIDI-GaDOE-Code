Global myValue As Variant
Global numberNewUsers As Integer
Sub FindNewUsers()

    LRowFirst = Worksheets(sheetNames(4)).Cells(Rows.Count, "A").End(xlUp).Row
    LRowSecond = Worksheets(sheetNames(5)).Cells(Rows.Count, "A").End(xlUp).Row
    
    Dim Rng1 As Range
    Dim Rng2 As Range

    Set Rng1 = Worksheets(sheetNames(4)).Range("A1:" + "A" + Right(Str(LRowFirst), Len(LRowFirst)))
    Set Rng2 = Worksheets(sheetNames(5)).Range("A1:" + "A" + Right(Str(LRowSecond), Len(LRowSecond)))

    Dim cellplaceholder As Range
    Dim finalString As String
    Dim counter As Integer
    finalString = "List of New Users: " + vbNewLine
    For Each cellplaceholder In Rng1
        If WorksheetFunction.CountIf(Rng2, cellplaceholder) = 0 Then
            finalString = finalString + "NAME: " + cellplaceholder.Value + " "
            finalString = finalString + "DISTRICT: " + cellplaceholder.Offset(0, 6) + vbNewLine
            counter = counter + 1
        End If
    Next
    If Len(finalString) = 0 Then
        finalString = "No new users"
    End If
    numberNewUsers = counter
    counter = 0
    finalString = finalString + "There are" + Str(numberNewUsers) + " new users registered in the portal"
    MsgBox finalString


End Sub
