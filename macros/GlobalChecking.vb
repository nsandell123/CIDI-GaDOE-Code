Option Explicit 
Global sheetNames(12) As String
Global mainWorkbook As Workbook
Sub GlobalChecking()
    Set mainWorkbook = ActiveWorkbook
    Dim i As Long
    For i = 1 To 12
        sheetNames(i) = mainWorkbook.Sheets(i).Name
    Next i

End Sub

