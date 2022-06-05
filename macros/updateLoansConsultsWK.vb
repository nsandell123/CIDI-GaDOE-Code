Sub updateLoansConsultsWK()
    Dim LastRowColumnA As Integer
    LastRowColumnA = Worksheets(13).Cells(Rows.Count, "A").End(xlUp).Row
    Dim updateAnchorCell As Range
    Set updateAnchorCell = Worksheets(13).Cells(LastRowColumnA + 1, "A")
    updateAnchorCell.Value = Left(inputDate, Len(inputDate) - 5) + " Totals "
    updateAnchorCell.Offset(0, 1).Value = updateAnchorCell.Offset(-1, 1).Value + totalLoans
    Worksheets(13).Range(updateAnchorCell, updateAnchorCell.Offset(0, 1)).Borders.LineStyle = xlContinuous
    Worksheets(13).Range(updateAnchorCell, updateAnchorCell.Offset(0, 1)).Borders.Weight = xlThin


    Dim lastRowColumnT As Integer
    lastRowColumnT = Worksheets(13).Cells(Rows.Count, "T").End(xlUp).Row
    Set updateAnchorCell = Worksheets(13).Cells(lastRowColumnT + 1, "T")
    updateAnchorCell.Value = Left(inputDate, Len(inputDate) - 5) + " Totals "
    updateAnchorCell.Offset(0, 1).Value = updateAnchorCell.Offset(-1, 1).Value + totalConsults
    Worksheets(13).Range(updateAnchorCell, updateAnchorCell.Offset(0, 1)).Borders.LineStyle = xlContinuous
    Worksheets(13).Range(updateAnchorCell, updateAnchorCell.Offset(0, 1)).Borders.Weight = xlThin

    MsgBox "Finished Updating Loans & Consults Weekly Sheet"
    
    
    
End Sub
