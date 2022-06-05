Sub updateUsersWeeklySheet()
    Dim LastRowColumnA As Integer
    LastRowColumnA = Worksheets(9).Cells(Rows.Count, "A").End(xlUp).Row
    Dim updateAnchorCell As Range
    Set updateAnchorCell = Worksheets(9).Cells(LastRowColumnA + 1, "A")
    updateAnchorCell.Value = inputDate
    updateAnchorCell.Offset(0, 1).Value = updateAnchorCell.Offset(-1, 1).Value + numberNewUsers
    MsgBox "Finished Updating Users Weekly Sheet"
End Sub
