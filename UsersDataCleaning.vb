Sub UsersDataCleaning()
    'Deleting Matthew Blake and Andrew Gelinas'
    For i = Cells(Rows.Count, 1).End(xlUp).Row To 1 Step -1
    If Cells(i, 1).Value = "Matthew Blake" Or Cells(i, 1).Value = "Andrew  Gelinas" Then
            Cells(i, 1).EntireRow.Delete
        End If
    Next i

    'Populating Oconee GLRS on 9 - Oconee'


 For i = Cells(Rows.Count, 1).End(xlUp).Row To 1 Step -1
    If Cells(i, 6).Value = "9 - Oconee" And Cells(i, 7).Value = "" Then
    Cells(i, 7).Value = "Oconee GLRS"
    End If
    Next i

End Sub
