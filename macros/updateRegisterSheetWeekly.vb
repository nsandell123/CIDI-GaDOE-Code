Sub updateRegisterSheetWeekly()
    Dim anchorCellCombined As Integer
    Dim anchorCellSoftware As Integer
    anchorCellCombined = Worksheets(sheetNames(1)).Range("K:K").Find(What:="Not Yet Prov", LookIn:=xlValues).Row - 6
    anchorCellSoftware = 103
    While Not (IsEmpty(Worksheets(7).Cells(anchorCellSoftware, "A")))
        anchorCellSoftware = anchorCellSoftware + 1
    Wend
    Rows(anchorCellSoftware).Insert
    
    Dim standardDate As Variant
    standardDate = Left(CStr(inputDate), 5)
    standardDate = standardDate + " Totals"
    Worksheets(7).Cells(anchorCellSoftware, "A").Value = standardDate
    Worksheets(7).Cells(anchorCellSoftware, "B").Value = Worksheets(1).Cells(anchorCellCombined, "L").Value
    Worksheets(7).Cells(anchorCellSoftware, "C").Value = Worksheets(1).Cells(anchorCellCombined, "M").Value
    Worksheets(7).Cells(anchorCellSoftware, "D").Value = Worksheets(1).Cells(anchorCellCombined, "N").Value
    Worksheets(7).Cells(anchorCellSoftware, "E").Value = Worksheets(1).Cells(anchorCellCombined, "O").Value
    Dim nameMonth As Variant
    nameMonth = MonthName(Left(inputDate, 2))
    nameMonth = Left(nameMonth, 3)
    Dim number As Variant
    number = Right(Left(inputDate, 5), 2)
    number = number + "-" + nameMonth
    anchorCellSoftware = 210
    While Not (IsEmpty(Worksheets(7).Cells(anchorCellSoftware, "A")))
        anchorCellSoftware = anchorCellSoftware + 1
    Wend
    Worksheets(sheetNames(7)).Cells(anchorCellSoftware, "A").Value = number
    Worksheets(sheetNames(7)).Cells(anchorCellSoftware, "B").Value = Worksheets(sheetNames(7)).Cells(anchorCellSoftware - 1, "B") + totalRequested
    Worksheets(sheetNames(7)).Cells(anchorCellSoftware, "C").Value = Worksheets(sheetNames(7)).Cells(anchorCellSoftware - 1, "C")
    anchorCellSoftware = 215
    While Not (IsEmpty(Worksheets(7).Cells(anchorCellSoftware, "F")))
        anchorCellSoftware = anchorCellSoftware + 1
    Wend
    
    Worksheets(sheetNames(7)).Cells(anchorCellSoftware - 1, "F").Value = number
    Worksheets(sheetNames(7)).Cells(anchorCellSoftware - 1, "G").Value = Worksheets(sheetNames(7)).Cells(anchorCellSoftware - 1, "G") + totalRegistered
    
    'Update Categories
    Dim j As Long
    Dim FindPortalMonth As Range
    Set FindPortalMonth = Worksheets(sheetNames(7)).Range("C:C").Find(What:="In Portal", LookIn:=xlValues)
    For j = 0 To 3
        Worksheets(sheetNames(7)).Cells(FindPortalMonth.Row + 1 + j, "C").Value = Worksheets(sheetNames(7)).Cells(FindPortalMonth.Row + 1 + j, "C").Value + totalCategories(j)
    Next j
    MsgBox "Finished Updating Register Sheet Weekly"            
End Sub
