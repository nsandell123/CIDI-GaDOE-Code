Sub UsersFormatting()
'
' UsersFormatting Macro
'

'
    Dim LRowFirst As Long
    LRowFirst = Worksheets(sheetNames(4)).Cells(Rows.Count, "D").End(xlUp).Row
    Dim Rng1 As Range
    Set Rng1 = Worksheets(sheetNames(4)).Range("D1:" + "D" + Right(Str(LRowFirst), Len(LRowFirst) - 1))
    
    Dim cellplaceholder As Range
    
    For Each cellplaceholder In Rng1
        cellplaceholder.NumberFormat = "m/d/yyyy h:mm"
        cellplaceholder.Value = cellplaceholder.Value
    Next cellplaceholder
    Worksheets(sheetNames(4)).Rows("1:1").Select
    Selection.Font.Bold = True
    Selection.RowHeight = 15
    Selection.ColumnWidth = 20
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
