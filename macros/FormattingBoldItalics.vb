Sub RequestsFormatting()
'
' RequestsFormatting Macro
'

'
    Dim LRowFirst As Long
    LRowFirst = Worksheets(sheetNames(2)).Cells(Rows.Count, "C").End(xlUp).Row
    Dim Rng1 As Range
    Set Rng1 = Worksheets(sheetNames(2)).Range("C1:" + "C" + Right(Str(LRowFirst), Len(LRowFirst) - 1))
    
    Dim cellplaceholder As Range
    
    For Each cellplaceholder In Rng1
        cellplaceholder.NumberFormat = "m/d/yyyy"
        cellplaceholder.Value = cellplaceholder.Value
    Next cellplaceholder
    
    Worksheets(sheetNames(2)).Select
    Worksheets(sheetNames(2)).Rows("1:1").Select
    Selection.Font.Bold = True
    Selection.RowHeight = 15
    Selection.ColumnWidth = 15
    'Removing Matthew Blake Records
    For i = Worksheets(sheetNames(2)).Cells(Rows.Count, 11).End(xlUp).Row To 1 Step -1
        If Worksheets(sheetNames(2)).Cells(i, 10).Value = "Blake" And Worksheets(sheetNames(2)).Cells(i, 11).Value = "Matthew" Then
            Worksheets(sheetNames(2)).Cells(i, 11).EntireRow.Delete
        End If
    Next i
End Sub
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
    
    Worksheets(sheetNames(4)).Select
    Worksheets(sheetNames(4)).Rows("1:1").Select
    Selection.Font.Bold = True
    Selection.RowHeight = 15
    Selection.ColumnWidth = 20
End Sub
