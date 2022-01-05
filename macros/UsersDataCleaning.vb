Sub UsersDataCleaning()
    Worksheets(sheetNames(4)).Select
    For i = Worksheets(sheetNames(4)).Cells(Rows.Count, 1).End(xlUp).Row To 1 Step -1
        'Deleting Matthew Blake and Andrew Gelinas'
        If Worksheets(sheetNames(4)).Cells(i, 1).Value = "Matthew Blake" Or Worksheets(sheetNames(4)).Cells(i, 1).Value = "Andrew  Gelinas" Then
            Worksheets(sheetNames(4)).Cells(i, 1).EntireRow.Delete
        End If
        'Populating Oconee GLRS on 9 - Oconee'
        If Worksheets(sheetNames(4)).Cells(i, 6).Value = "9 - Oconee" And Worksheets(sheetNames(4)).Cells(i, 7).Value = "" Then
            Worksheets(sheetNames(4)).Cells(i, 7).Value = "Oconee GLRS"
        End If
    Next i
    
End Sub
