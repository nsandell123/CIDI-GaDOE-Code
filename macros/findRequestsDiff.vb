Option Explicit
Global rowRequestNumbers() As Long
Global numberNewRequests As Long
Global totalLoans As Integer
Global totalConsults As Integer
Sub findRequestsDiff()
    Dim numbers() As Long
    Dim numberLoans As Integer
    Dim numberConsults As Integer
    numberLoans = 0
    numberConsults = 0
    Dim counter As Long
    counter = 0
    Dim request As Variant
    Dim LRowFirst As Long
    Dim LRowSecond As Long
    LRowFirst = Worksheets(sheetNames(2)).Cells(Rows.Count, "C").End(xlUp).Row
    LRowSecond = Worksheets(sheetNames(3)).Cells(Rows.Count, "C").End(xlUp).Row
    'This next section is dedicated to finding the requests diff in terms of the row number
    Dim Rng1 As Range
    Dim Rng2 As Range
    Set Rng1 = Worksheets(sheetNames(2)).Range("C1:" + "C" + Right(Str(LRowFirst), Len(LRowFirst) - 1))
    Set Rng2 = Worksheets(sheetNames(3)).Range("C1:" + "C" + Right(Str(LRowSecond), Len(LRowSecond) - 1))

    Dim cellplaceholder As Range
    Dim finalString As String
    For Each cellplaceholder In Rng1
        If WorksheetFunction.CountIf(Rng2, cellplaceholder) = 0 Then
            ReDim Preserve numbers(counter)
            numbers(counter) = cellplaceholder.Row
            counter = counter + 1
        End If
    Next cellplaceholder
    If counter = 0 Then
        MsgBox "There were no new requests"
        Exit Sub
    End If
    'Now I will focus on building the finalString
    finalString = finalString + "NEW REQUESTS: " + vbNewLine
    For Each request In numbers
        Dim firstName As String
        Dim lastName As String
        Dim District As String
        Dim requestDate As String
        Dim serviceRequested As String
        firstName = Worksheets(2).Cells(request, "K").Value
        lastName = Worksheets(2).Cells(request, "J").Value
        District = Worksheets(2).Cells(request, "M").Value
        requestDate = Worksheets(2).Cells(request, "C").Value
        serviceRequested = Worksheets(2).Cells(request, "B").Value
        finalString = finalString + firstName + " " + lastName + " from District " + District + " requested " + serviceRequested + " on " + requestDate + vbNewLine
        If serviceRequested = "Software" Then
            Dim numberOfSoftwareItems As Integer
            Dim softwareTools As String
            numberOfSoftwareItems = Worksheets(2).Cells(request, "H").Value
            softwareTools = Worksheets(2).Cells(request, "G").Value
            finalString = finalString + " He/She requested" + Str(numberOfSoftwareItems) + " " + softwareTools + " each." + vbNewLine
        End If
        If serviceRequested = "Product Loan" Then
            finalString = finalString + " He/She requested " + Worksheets(2).Cells(request, "I") + vbNewLine
            numberLoans = numberLoans + 1
        End If
        If serviceRequested = "Consulting" Then
            numberConsults = numberConsults + 1
        End If
        
    Next request
    totalConsults = numberConsults
    totalLoans = numberLoans
    numberNewRequests = counter
    counter = 0
    rowRequestNumbers = numbers
    finalString = finalString + "There were a total of" + Str(numberNewRequests) + " requests"
    MsgBox finalString
End Sub
