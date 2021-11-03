Sub RequestsFormatting()
'
' RequestsFormatting Macro
'

'
    Rows("1:1").Select
    Selection.Font.Bold = True
    Cells.Select
    Selection.RowHeight = 15
    Selection.ColumnWidth = 15
End Sub
Sub UsersFormatting()
'
' UsersFormatting Macro
'

'
    Rows("1:1").Select
    Selection.Font.Bold = True
    Cells.Select
    Selection.RowHeight = 15
    Selection.ColumnWidth = 20
    Range("E4").Select
End Sub
