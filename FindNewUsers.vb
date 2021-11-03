Sub FindNewUsers()

    Dim mainworkBook As Workbook
    Dim firstSheet As String
    Dim secondSheet As String
    Dim flag As Boolean

    Set mainworkBook = ActiveWorkbook
    For i = 1 To mainworkBook.Sheets.Count
        If InStr(LCase(mainworkBook.Sheets(i).Name), "user") <> 0 And InStr(LCase(mainworkBook.Sheets(i).Name), "ga") <> 0 And Not (flag) Then
           firstSheet = mainworkBook.Sheets(i).Name
           flag = True

        End If
        If InStr(LCase(mainworkBook.Sheets(i).Name), "user") <> 0 And InStr(LCase(mainworkBook.Sheets(i).Name), "ga") <> 0 And (flag) Then
            secondSheet = mainworkBook.Sheets(i).Name

        End If
    Next i
    LRowFirst = Worksheets(firstSheet).Cells(Rows.Count, "A").End(xlUp).Row
    LRowSecond = Worksheets(secondSheet).Cells(Rows.Count, "A").End(xlUp).Row

    Dim Rng1 As Range
    Dim Rng2 As Range

    Set Rng1 = Worksheets(firstSheet).Range("A1:" + "A" + Right(Str(LRowFirst), Len(LRowFirst)))
    Set Rng2 = Worksheets(secondSheet).Range("A1:" + "A" + Right(Str(LRowSecond), Len(LRowSecond)))

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

    Dim userSheet1 As String
    Dim userSheet2 As String
    For i = 1 To mainworkBook.Sheets.Count
        If InStr(LCase(mainworkBook.Sheets(i).Name), "users") <> 0 And InStr(LCase(mainworkBook.Sheets(i).Name), "weekly") <> 0 Then
           userSheet1 = mainworkBook.Sheets(i).Name

        End If
        If InStr(LCase(mainworkBook.Sheets(i).Name), "user") <> 0 And InStr(LCase(mainworkBook.Sheets(i).Name), "monthly") <> 0 Then
            userSheet2 = mainworkBook.Sheets(i).Name

        End If
    Next i


    LastRowWeekly = Worksheets(userSheet1).Cells(Rows.Count, "B").End(xlUp).Row
    LastRowMonthly = Worksheets(userSheet2).Cells(Rows.Count, "B").End(xlUp).Row

    Dim myValue As Variant
    myValue = InputBox("What is today's date? (MM/DD/YYYY)")

    Worksheets(userSheet2).Cells(LastRowMonthly, "B") = Str(CLng(Worksheets(userSheet2).Cells(LastRowMonthly, "B")) + counter)

    Worksheets(userSheet1).Cells(LastRowWeekly + 1, "B") = Worksheets(userSheet1).Cells(LastRowWeekly, "B") + counter

    Worksheets(userSheet1).Cells(LastRowWeekly + 1, "A") = myValue
    Worksheets(userSheet1).Cells(LastRowWeekly + 1, "A").HorizontalAlignment = xlRight


    MsgBox finalString







End Sub
