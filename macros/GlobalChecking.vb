Global mainWorkbook As Workbook
Global inputDate As Variant
Sub GlobalChecking()
    inputDate = InputBox("What is today's date? (MM/DD/YYYY)")
    Dim LastRowTest As Integer
    Set mainWorkbook = ActiveWorkbook
    If mainWorkbook.Sheets.Count <> 13 Then
        MsgBox "There should be 13 sheets in the workbook"
        MsgBox "Please run Initialization Macro Again"
	Exit Sub
    End If

    'Checking if Sheets are in correct order
    Dim regexSheetOne As Object
    Set regexSheetOne = New RegExp
    regexSheetOne.Pattern = ("\bCombined Data \d\d\d\d\d\d\b")
    If regexSheetOne.Test(Worksheets(2).Name) <> True Then
        MsgBox "Sheet 2 must be in the format Combined Data MMDDYY"
        MsgBox "Please run Initialization Macro Again"
        Exit Sub
    End If


    Dim regexSheetTwo As Object
    Set regexSheetTwo = New RegExp
    regexSheetTwo.Pattern = ("\bGA DoE Requests \d\d\d\d\d\d\b")
    If regexSheetTwo.Test(Worksheets(3).Name) <> True Then
        MsgBox "Sheet 3 must be in the format Combined Data GA DoE Requests MMDDYY"
        MsgBox "Please run Initialization Macro Again"
        Exit Sub
    End If
    If regexSheetTwo.Test(Worksheets(4).Name) <> True Then
        MsgBox "Sheet 4 must be in the format Combined Data GA DoE Requests MMDDYY"
        MsgBox "Please run Initialization Macro Again"
        Exit Sub
    End If

    Dim regexSheetThree As Object
    Set regexSheetThree = New RegExp
    regexSheetThree.Pattern = ("\bGA DoE Users \d\d\d\d\d\d\b")
    If regexSheetThree.Test(Worksheets(5).Name) <> True Then
        MsgBox "Sheet 5 must be in the format Combined Data GA DoE Users MMDDYY"
        MsgBox "Please run Initialization Macro Again"
        Exit Sub
    End If
    If regexSheetThree.Test(Worksheets(6).Name) <> True Then
        MsgBox "Sheet 6 must be in the format Combined Data GA DoE Users MMDDYY"
        MsgBox "Please run Initialization Macro Again"
        Exit Sub
    End If

    Dim regexSheetFour As Object
    Set regexSheetFour = New RegExp
    regexSheetFour.Pattern = ("\bRegistr-Softw Mo\b")
    If regexSheetFour.Test(Worksheets(7).Name) <> True Then
        MsgBox "Sheet 7 must say Regstr-Softw Mo"
        MsgBox "Please run Initialization Macro Again"
        Exit Sub
    End If

    Dim regexSheetFive As Object
    Set regexSheetFive = New RegExp
    regexSheetFive.Pattern = ("\bRegistr-Softw WK\b")
    If regexSheetFive.Test(Worksheets(8).Name) <> True Then
        MsgBox "Sheet 8 must say Registr-Softw WK"
        MsgBox "Please run Initialization Macro Again"
        Exit Sub
    End If
    Dim regexSheetSix As Object
    Set regexSheetSix = New RegExp
    regexSheetSix.Pattern = ("\bUsers Weekly\b")
    If regexSheetSix.Test(Worksheets(9).Name) <> True Then
        MsgBox "Sheet 9 must say Users Weekly"
        MsgBox "Please run Initialization Macro Again"
        Exit Sub
    End If

    Dim regexSheetSeven As Object
    Set regexSheetSeven = New RegExp
    regexSheetSeven.Pattern = ("\bUsers Monthly\b")
    If regexSheetSeven.Test(Worksheets(10).Name) <> True Then
        MsgBox "Sheet 10 should say  Users Monthly"
        MsgBox "Please run Initialization Macro Again"
        Exit Sub
    End If
    Dim regexSheetEight As Object
    Set regexSheetEight = New RegExp
    regexSheetEight.Pattern = ("\bLoans & Consults\b")
    If regexSheetEight.Test(Worksheets(11).Name) <> True Then
        MsgBox "Sheet 11 should say Loans & Consults"
        MsgBox "Please run Initialization Macro Again"
        Exit Sub
    End If

    Dim regexSheetNine As Object
    Set regexSheetNine = New RegExp
    regexSheetNine.Pattern = ("\bLoans-Consults Mo\b")
    If regexSheetNine.Test(Worksheets(12).Name) <> True Then
        MsgBox "Sheet 12 should say Loans-Consults Mo"
        MsgBox "Please run Initialization Macro Again"
        Exit Sub
    End If

    Dim regexSheetTen As Object
    Set regexSheetTen = New RegExp
    regexSheetTen.Pattern = ("\bLoans-Consults WK\b")
    If regexSheetTen.Test(Worksheets(13).Name) <> True Then
        MsgBox "Sheet 13 should say Loans-Consults WK"
        MsgBox "Please run Initialization Macro Again"
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
    While Not (IsEmpty(Worksheets(10).Cells(2 + counter, "A")))
        If Not (Dict.Exists(Worksheets(10).Cells(2 + counter, "A").Value)) Then
            MsgBox "In Worksheet 10, Cell A" + Right(Str(2 + counter), Len(Str(2 + counter)) - 1) + "should be one of Jan, Feb, Mar, Apr, May, June, July, Aug, Sept, Oct, Nov, Dec "
            MsgBox "Be sure to run this module again to (maybe) find other months which are not formatted correctly"
            Exit Sub
        End If
        counter = counter + 1
    Wend
    
    counter = 0
    While Not (IsEmpty(Worksheets(12).Cells(2 + counter, "A")))
        If Not (Dict.Exists(Worksheets(12).Cells(2 + counter, "A").Value)) Then
            MsgBox "In Worksheet 12, Cell A" + Right(Str(2 + counter), Len(Str(2 + counter)) - 1) + "should be one of Jan, Feb, Mar, Apr, May, June, July, Aug, Sept, Oct, Nov, Dec "
            MsgBox "Be sure to run this module again to (maybe) find other months which are not formatted correctly"
            Exit Sub
        End If
        counter = counter + 1
    Wend
    counter = 0
    While Not (IsEmpty(Worksheets(12).Cells(2 + counter, "E")))
        If Not (Dict.Exists(Worksheets(12).Cells(2 + counter, "E").Value)) Then
            MsgBox "In Worksheet 12, Cell E" + Right(Str(2 + counter), Len(Str(2 + counter)) - 1) + "should be one of Jan, Feb, Mar, Apr, May, June, July, Aug, Sept, Oct, Nov, Dec "
            MsgBox "Be sure to run this module again to (maybe) find other months which are not formatted correctly"
            Exit Sub
        End If
        counter = counter + 1
    Wend
    
    Dim anchorCellRequested As Integer

    anchorCellRequested = Worksheets(7).Range("B:B").Find(What:="Requested Software", LookIn:=xlValues).Row + 1
    While Not (IsEmpty(Worksheets(7).Cells(2 + anchorCellRequested, "A")))
        If Not (Dict.Exists(Worksheets(7).Cells(2 + anchorCellRequested, "A").Value)) Then
            MsgBox "In Worksheet 7, Cell A" + Right(Str(2 + anchorCellRequested), Len(Str(2 + anchorCellRequested)) - 1) + "should be one of Jan, Feb, Mar, Apr, May, June, July, Aug, Sept, Oct, Nov, Dec "
            MsgBox "Be sure to run this module again to (maybe) find other months which are not formatted correctly"
            Exit Sub
        End If
        anchorCellRequested = anchorCellRequested + 1
    Wend
    
    Dim anchorCellRegistered As Integer
    anchorCellRegistered = Worksheets(7).Range("G:G").Find(What:="Register in Portal", LookIn:=xlValues).Row + 1
    While Not (IsEmpty(Worksheets(7).Cells(2 + anchorCellRegistered, "F")))
        If Not (Dict.Exists(Worksheets(7).Cells(2 + anchorCellRegistered, "F").Value)) Then
            MsgBox "In Worksheet 7, Cell F" + Right(Str(2 + anchorCellRegistered), Len(Str(2 + anchorCellRegistered)) - 1) + "should be one of Jan, Feb, Mar, Apr, May, June, July, Aug, Sept, Oct, Nov, Dec "
            MsgBox "Be sure to run this module again to (maybe) find other months which are not formatted correctly"
            Exit Sub
        End If
        anchorCellRegistered = anchorCellRegistered + 1
    Wend
    
    MsgBox "All Tests Passed"
End Sub
