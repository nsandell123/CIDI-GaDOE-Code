Option Explicit
Global sheetNames(12) As String
Global mainWorkbook As Workbook
Global inputDate As Variant
Sub GlobalChecking()
    inputDate = InputBox("What is today's date? (MM/DD/YYYY)")
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
    
    'Checking if Sheets are in correct order
    Dim regexSheetOne As Object
    Set regexSheetOne = New RegExp
    regexSheetOne.Pattern = ("\bCombined Data \d\d\d\d\d\d\b")
    If regexSheetOne.Test(sheetNames(1)) <> True Then
        MsgBox "Sheet 1 must be in the format Combined Data MMDDYY"
        Exit Sub
    End If

    
    Dim regexSheetTwo As Object
    Set regexSheetTwo = New RegExp
    regexSheetTwo.Pattern = ("\bGA DoE Requests \d\d\d\d\d\d\b")
    If regexSheetTwo.Test(sheetNames(2)) <> True Then
        MsgBox "Sheet 2 must be in the format Combined Data GA DoE Requests MMDDYY"
        Exit Sub
    End If
    If regexSheetTwo.Test(sheetNames(3)) <> True Then
        MsgBox "Sheet 3 must be in the format Combined Data GA DoE Requests MMDDYY"
        Exit Sub
    End If
    
    Dim regexSheetThree As Object
    Set regexSheetThree = New RegExp
    regexSheetThree.Pattern = ("\bGA DoE Users \d\d\d\d\d\d\b")
    If regexSheetThree.Test(sheetNames(4)) <> True Then
        MsgBox "Sheet 4 must be in the format Combined Data GA DoE Users MMDDYY"
        Exit Sub
    End If
    If regexSheetThree.Test(sheetNames(5)) <> True Then
        MsgBox "Sheet 5 must be in the format Combined Data GA DoE Users MMDDYY"
        Exit Sub
    End If
    
    Dim regexSheetFour As Object
    Set regexSheetFour = New RegExp
    regexSheetFour.Pattern = ("\bRegistr-Softw Mo\b")
    If regexSheetFour.Test(sheetNames(6)) <> True Then
        MsgBox "Sheet 6 must say Regstr-Softw Mo"
        Exit Sub
    End If
    
    Dim regexSheetFive As Object
    Set regexSheetFive = New RegExp
    regexSheetFive.Pattern = ("\bRegistr-Softw WK\b")
    If regexSheetFive.Test(sheetNames(7)) <> True Then
        MsgBox "Sheet 7 must say Registr-Softw WK"
        Exit Sub
    End If
    Dim regexSheetSix As Object
    Set regexSheetSix = New RegExp
    regexSheetSix.Pattern = ("\bUsers Weekly\b")
    If regexSheetSix.Test(sheetNames(8)) <> True Then
        MsgBox "Sheet 8 must say Users Weekly"
        Exit Sub
    End If
    
    Dim regexSheetSeven As Object
    Set regexSheetSeven = New RegExp
    regexSheetSeven.Pattern = ("\bUsers Monthly\b")
    If regexSheetSeven.Test(sheetNames(9)) <> True Then
        MsgBox "Sheet 9 should say  Users Monthly"
        Exit Sub
    End If
    Dim regexSheetEight As Object
    Set regexSheetEight = New RegExp
    regexSheetEight.Pattern = ("\bLoans & Consults\b")
    If regexSheetEight.Test(sheetNames(10)) <> True Then
        MsgBox "Sheet 10 should say Loans & Consults"
        Exit Sub
    End If
    
    Dim regexSheetNine As Object
    Set regexSheetNine = New RegExp
    regexSheetNine.Pattern = ("\bLoans-Consults Mo\b")
    If regexSheetNine.Test(sheetNames(11)) <> True Then
        MsgBox "Sheet 11 should say Loans-Consults Mo"
        Exit Sub
    End If
    
    Dim regexSheetTen As Object
    Set regexSheetTen = New RegExp
    regexSheetTen.Pattern = ("\bLoans-Consults WK\b")
    If regexSheetTen.Test(sheetNames(12)) <> True Then
        MsgBox "Sheet 11 should say Loans-Consults WK"
        Exit Sub
    End If
    
    'Check if all months in Users Monthly have the valid abbreviation
    Dim Dict As New Scripting.Dictionary
    Dict.Add "Jan", "1"
    Dict.Add "Feb", "2"
    Dict.Add "Mar", "3"
    Dict.Add "Apr", "4"
    Dict.Add "May", "5"
    Dict.Add "June", "6"
    Dict.Add "July", "7"
    Dict.Add "Aug", "8"
    Dict.Add "Sept", "9"
    Dict.Add "Oct", "10"
    Dict.Add "Nov", "11"
    Dict.Add "Dec", "12"
    
    Dim counter As Integer
    counter = 0
    While Not (IsEmpty(Worksheets(9).Cells(2 + counter, "A")))
        If Not (Dict.Exists(Worksheets(9).Cells(2 + counter, "A").Value)) Then
            MsgBox "In Worksheet 9, Cell A" + Right(Str(2 + counter), Len(Str(2 + counter)) - 1) + "should be one of Jan, Feb, Mar, Apr, May, June, July, Aug, Sept, Oct, Nov, Dec "
            MsgBox "Be sure to run this module again to (maybe) find other months which are not formatted correctly"
            Exit Sub
        End If
        counter = counter + 1
    Wend
    
    
    
End Sub
