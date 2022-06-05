Sub RequestsFormatting()
'
' RequestsFormatting Macro
'

'
    Dim LRowFirst As Long
    LRowFirst = Worksheets(3).Cells(Rows.Count, "C").End(xlUp).Row
    Dim Rng1 As Range
    Set Rng1 = Worksheets(3).Range("C1:" + "C" + Right(Str(LRowFirst), Len(LRowFirst) - 1))
    
    Dim cellplaceholder As Range
    
    For Each cellplaceholder In Rng1
        cellplaceholder.NumberFormat = "m/d/yyyy"
        cellplaceholder.Value = cellplaceholder.Value
    Next cellplaceholder

    Worksheets(3).Select
    Worksheets(3).Rows("1:1").Select
    Selection.Font.Bold = True
    Selection.RowHeight = 15
    Selection.ColumnWidth = 15
    'Removing Matthew Blake Records
    For i = Worksheets(3).Cells(Rows.Count, 11).End(xlUp).Row To 1 Step -1
        If Worksheets(3).Cells(i, 10).Value = "Blake" And Worksheets(3).Cells(i, 11).Value = "Matthew" Then
            Worksheets(3).Cells(i, 11).EntireRow.Delete
        End If
    Next i
End Sub
