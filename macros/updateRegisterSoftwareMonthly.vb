Sub updateRegisterMonthly()
    'Update Software Numbers
    Dim anchorCellCombined As Integer
    Dim anchorCellSoftware As Integer
    anchorCellCombined = Worksheets(2).Range("K:K").Find(What:="Not Yet Prov", LookIn:=xlValues).Row - 7
    anchorCellSoftware = Worksheets(7).Range("A:A").Find(What:="New", LookIn:=xlValues).Row

    Worksheets(7).Cells(anchorCellSoftware, "B").Value = Worksheets(2).Cells(anchorCellCombined, "L").Value
    Worksheets(7).Cells(anchorCellSoftware, "C").Value = Worksheets(2).Cells(anchorCellCombined, "M").Value
    Worksheets(7).Cells(anchorCellSoftware, "D").Value = Worksheets(2).Cells(anchorCellCombined, "N").Value
    Worksheets(7).Cells(anchorCellSoftware, "E").Value = Worksheets(2).Cells(anchorCellCombined, "O").Value
    'Update Requested District Numbers
    Dim anchorCellRequested As Integer
    anchorCellRequested = Worksheets(7).Range("B:B").Find(What:="Requested Software", LookIn:=xlValues).Row + 1
    While Not (IsEmpty(Worksheets(7).Cells(2 + anchorCellRequested, "A")))
        anchorCellRequested = anchorCellRequested + 1
    Wend
    Rows(anchorCellRequested + 2).Insert
    Dim monthOnSheet As String
    monthOnSheet = Worksheets(7).Cells(anchorCellRequested + 1, "A").Value
    Dim c As Collection
    Set c = New Collection
    c.Add "Jan", "1"
    c.Add "Feb", "2"
    c.Add "Mar", "3"
    c.Add "Apr", "4"
    c.Add "May", "5"
    c.Add "June", "6"
    c.Add "July", "7"
    c.Add "Aug", "8"
    c.Add "Sept", "9"
    c.Add "Oct", "10"
    c.Add "Nov", "11"
    c.Add "Dec", "12"
    Dim actualMonthNumber As String
    actualMonthNumber = Right(Str(CInt(Left(inputDate, 2))), 1)
    Dim mappedActualMonthNumber As String
    mappedActualMonthNumber = c(actualMonthNumber)
    If mappedActualMonthNumber <> Worksheets(7).Cells(anchorCellRequested + 1, "A").Value Then
        mappedActualMonthNumber = c(actualMonthNumber)
        Worksheets(7).Cells(anchorCellRequested + 1, "A").Offset(1, 0).Value = mappedActualMonthNumber
        Worksheets(7).Cells(anchorCellRequested + 1, "A").Offset(1, 1).Value = CInt(Worksheets(7).Cells(anchorCellRequested + 1, "A").Offset(0, 1)) + totalRequested
        Worksheets(7).Cells(anchorCellRequested + 1, "A").Offset(1, 2).Value = 245
    Else
        Worksheets(7).Cells(anchorCellRequested + 1, "A").Offset(0, 1).Value = CInt(Worksheets(7).Cells(anchorCellRequested + 1, "A").Offset(0, 1)) + totalRequested
        Worksheets(7).Cells(anchorCellRequested + 1, "A").Offset(0, 2).Value = 245
    End If
    'Update Registered Districts Numbers
    Dim anchorCellRegistered As Integer
    anchorCellRegistered = Worksheets(7).Range("G:G").Find(What:="Register In Portal", LookIn:=xlValues).Row + 1
    While Not (IsEmpty(Worksheets(7).Cells(2 + anchorCellRegistered, "F")))
        anchorCellRegistered = anchorCellRegistered + 1
    Wend
    monthOnSheet = Worksheets(7).Cells(anchorCellRegistered + 1, "F").Value
    If mappedActualMonthNumber <> Worksheets(7).Cells(anchorCellRegistered + 1, "F").Value Then
        mappedActualMonthNumber = c(actualMonthNumber)
        Worksheets(7).Cells(anchorCellRegistered + 1, "F").Offset(1, 0).Value = mappedActualMonthNumber
        Worksheets(7).Cells(anchorCellRegistered + 1, "F").Offset(1, 1).Value = CInt(Worksheets(7).Cells(anchorCellRegistered + 1, "F").Offset(0, 1)) + totalRegistered
    Else
        Worksheets(7).Cells(anchorCellRegistered + 1, "F").Offset(0, 1).Value = CInt(Worksheets(7).Cells(anchorCellRegistered + 1, "F").Offset(0, 1)) + totalRegistered
    End If
    'Update Categories
    Dim j As Long
    Set FindPortalMonth = Worksheets(7).Range("C:C").Find(What:="In Portal", LookIn:=xlValues)
    For j = 0 To 3
        Worksheets(7).Cells(FindPortalMonth.Row + 1 + j, "C").Value = Worksheets(7).Cells(FindPortalMonth.Row + 1 + j, "C").Value + totalCategories(j)
    Next j
    MsgBox "Finished Updating Register Software Monthly"

End Sub

