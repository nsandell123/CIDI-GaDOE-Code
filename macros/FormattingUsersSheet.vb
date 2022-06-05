Sub UsersFormatting()
'
' UsersFormatting Macro
'

'
    Dim LRowFirst As Long
    LRowFirst = Worksheets(5).Cells(Rows.Count, "D").End(xlUp).Row
    Dim Rng1 As Range
    Set Rng1 = Worksheets(5).Range("D1:" + "D" + Right(Str(LRowFirst), Len(LRowFirst) - 1))
    
    Dim cellplaceholder As Range
    
    For Each cellplaceholder In Rng1
        cellplaceholder.NumberFormat = "m/d/yyyy h:mm"
        cellplaceholder.Value = cellplaceholder.Value
    Next cellplaceholder
    Worksheets(5).Rows(1).Font.Bold = True
    Worksheets(5).Rows(1).RowHeight = 15
    Worksheets(5).Rows(1).ColumnWidth = 20
    For i = Worksheets(5).Cells(Rows.Count, 1).End(xlUp).Row To 1 Step -1
        'Deleting Matthew Blake and Andrew Gelinas'
        If Worksheets(5).Cells(i, 1).Value = "Matthew Blake" Or Worksheets(5).Cells(i, 1).Value = "Andrew  Gelinas" Then
            Worksheets(5).Cells(i, 1).EntireRow.Delete
        End If
        'Populating Oconee GLRS on 9 - Oconee'
        If Worksheets(5).Cells(i, 6).Value = "9 - Oconee" And Worksheets(5).Cells(i, 7).Value = "" Then
            Worksheets(5).Cells(i, 7).Value = "Oconee GLRS"
        End If
    Next i
End Sub
