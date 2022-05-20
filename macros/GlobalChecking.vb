Sub GlobalChecking()
    Set mainWorkbook = ActiveWorkbook
    If mainWorkbook.Sheets.Count <> 12 Then
        MsgBox "There should be 12 sheets in the workbook"
    End If
    Dim i As Long
    For i = 1 To 12
        sheetNames(i) = mainWorkbook.Sheets(i).Name
    Next i
End Sub

