Sub RequestsFormatting()
'
' RequestsFormatting Macro
'

'

    Worksheets(sheetNames(2)).Select
    Worksheets(sheetNames(2)).Rows("1:1").Select
    Selection.Font.Bold = True
    Selection.RowHeight = 15
    Selection.ColumnWidth = 15
End Sub
Sub UsersFormatting()
'
' UsersFormatting Macro
'

'
    Worksheets(sheetNames(5)).Select
    Worksheets(sheetNames(5)).Rows("1:1").Select
    Selection.Font.Bold = True
    Selection.RowHeight = 15
    Selection.ColumnWidth = 20
End Sub
