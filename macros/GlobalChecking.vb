Option Explicit
Global sheetNames(12) As String
Global mainWorkbook As Workbook
Sub GlobalChecking()
    Dim LastRowTest As Integer
    Set mainWorkbook = ActiveWorkbook
    If mainWorkbook.Sheets.Count <> 12 Then
        MsgBox "There should be 12 sheets in the workbook"
    End If
    Dim i As Long
    For i = 1 To 12
        sheetNames(i) = Worksheets(i).Name
    Next i
    
    Dim LastRowDistricts As Long
    LastRowDistricts = Worksheets(1).Cells(Rows.Count, "E").End(xlUp).Row
    Dim alphabeticalIndex As Integer
    alphabeticalIndex = 2
    For i = 2 To LastRowDistricts - 1
        If IsEmpty(Cells(i, "E")) Or IsEmpty(Cells(i + 1, "E")) Then
            alphabeticalIndex = alphabeticalIndex + 1
        Else
            If StrComp(Cells(i, "E").Value, Cells(i + 1, "E").Value, vbTextCompare) = -1 Or StrComp(Cells(i, "E").Value, Cells(i + 1, "E").Value, vbTextCompare) = 0 Then
                alphabeticalIndex = alphabeticalIndex + 1
            Else
                Exit For
            End If
        End If
    Next
    
    If alphabeticalIndex <> LastRowDistricts Then
        MsgBox "The districts on Sheet 1 are not in alphabetical order" & Str(alphabeticalIndex) & " is out of order"
    End If
    
    
    

    
End Sub

