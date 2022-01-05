Sub FindNewUsers()

    LRowFirst = Worksheets(sheetNames(4)).Cells(Rows.Count, "A").End(xlUp).Row
    LRowSecond = Worksheets(sheetNames(5)).Cells(Rows.Count, "A").End(xlUp).Row
    
    Dim Rng1 As Range
    Dim Rng2 As Range

    Set Rng1 = Worksheets(sheetNames(4)).Range("A1:" + "A" + Right(Str(LRowFirst), Len(LRowFirst)))
    Set Rng2 = Worksheets(sheetNames(5)).Range("A1:" + "A" + Right(Str(LRowSecond), Len(LRowSecond)))

    Dim cellplaceholder As Range
    Dim counter As Long
    Dim finalString As String
    For Each cellplaceholder In Rng1
        If WorksheetFunction.CountIf(Rng2, cellplaceholder) = 0 Then
            finalString = finalString + cellplaceholder.Value + vbNewLine
            counter = counter + 1
        End If
    Next
    If Len(finalString) = 0 Then
        finalString = "No new users"
    End If

    MsgBox finalString
    


    LastRowWeekly = Worksheets(sheetNames(8)).Cells(Rows.Count, "B").End(xlUp).Row
    LastRowMonthly = Worksheets(sheetNames(9)).Cells(Rows.Count, "B").End(xlUp).Row

    Dim myValue As Variant
    myValue = InputBox("What is today's date? (MM/DD/YYYY)")
    
    Worksheets(sheetNames(9)).Cells(LastRowMonthly, "B") = Str(CLng(Worksheets(sheetNames(9)).Cells(LastRowMonthly, "B")) + counter)

    Worksheets(sheetNames(8)).Cells(LastRowWeekly + 1, "B") = Worksheets(sheetNames(8)).Cells(LastRowWeekly, "B") + counter
    Worksheets(sheetNames(8)).Cells(LastRowWeekly + 1, "A") = myValue
    Worksheets(sheetNames(8)).Cells(LastRowWeekly + 1, "A").HorizontalAlignment = xlRight


    







End Sub
