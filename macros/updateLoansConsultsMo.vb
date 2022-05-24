Sub updateLoansConsultsMo()
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
    
    Dim LastRowColumnA As Integer
    LastRowColumnA = Worksheets(11).Cells(Rows.Count, "A").End(xlUp).Row
    Dim anchorCell As Range
    Set anchorCell = Worksheets(11).Cells(LastRowColumnA, "A")
    Dim actualMonthNumber As String
    actualMonthNumber = Right(Str(CInt(Left(inputDate, 2))), 1)
    Dim mappedActualMonthNumber As String
    mappedActualMonthNumber = c(actualMonthNumber)
    If mappedActualMonthNumber <> anchorCell.Value Then
        mappedActualMonthNumber = c(actualMonthNumber)
        anchorCell.Offset(1, 0).Value = mappedActualMonthNumber
        anchorCell.Offset(1, 1).Value = CInt(anchorCell.Offset(0, 1)) + totalLoans
    Else
        anchorCell.Offset(0, 1).Value = CInt(anchorCell.Offset(0, 1)) + totalLoans
    End If
    
    Dim LastRowColumnE As Integer
    LastRowColumnE = Worksheets(11).Cells(Rows.Count, "E").End(xlUp).Row
    Dim anchorCellE As Range
    Set anchorCellE = Worksheets(11).Cells(LastRowColumnE, "E")
    Dim actualMonthNumberE As String
    actualMonthNumberE = Right(Str(CInt(Left(inputDate, 2))), 1)
    Dim mappedActualMonthNumberE As String
    mappedActualMonthNumberE = c(actualMonthNumberE)
    If mappedActualMonthNumberE <> anchorCellE.Value Then
        mappedActualMonthNumberE = c(actualMonthNumberE)
        anchorCellE.Offset(1, 0).Value = mappedActualMonthNumberE
        anchorCellE.Offset(1, 1).Value = CInt(anchorCellE.Offset(0, 1)) + totalConsults
    Else
        anchorCellE.Offset(0, 1).Value = CInt(anchorCellE.Offset(0, 1)) + totalConsults
    End If
    
    MsgBox "Loans & Consults MO has finished updating"
End Sub
