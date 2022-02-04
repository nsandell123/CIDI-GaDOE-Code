Option Explicit
Sub findRequestsDiff()
    Dim numbers() As Long
    ReDim numbers(5)
    Dim counter As Long
    counter = 0
    Dim i As Integer
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
            numbers(counter) = cellplaceholder.Row
            counter = counter + 1
            'If we've run out of space in the numbers array, we have to resize using the ReDim function
            If UBound(numbers) - LBound(numbers) + 1 < counter Then
                ReDim Preserve numbers(counter)
            End If
            finalString = finalString + Str(cellplaceholder.Value) + vbNewLine
        End If
    Next cellplaceholder
    If Len(finalString) = 0 Then
        finalString = "No new requests"
    End If
    MsgBox finalString

    'This next part will be dedicated to finding the location in alphabetical order
    Dim numberIndex As Integer
    Dim numberProduct As Integer
    Dim numberConsult As Integer
    For numberIndex = 0 To counter - 1
        Dim requestsPointer As Integer
        requestsPointer = numbers(numberIndex)
        Dim District As String
        District = Worksheets(sheetNames(2)).Cells(requestsPointer, "M").Value
        Dim lastRowE As Integer
        lastRowE = Worksheets(1).Cells(Rows.Count, "E").End(xlUp).Row
        Dim position As Integer
        position = 2
        While position <= lastRowE And StrComp(Worksheets(1).Cells(position, "E").Value, District) <> 1
            position = position + 1
        Wend
        Worksheets(1).Rows(position).Insert
        Worksheets(1).Cells(position, "G").ClearFormats
        Worksheets(1).Cells(position, "A") = "Portal"
        Worksheets(1).Cells(position, "B") = Worksheets(sheetNames(2)).Cells(requestsPointer, "J")
        Worksheets(1).Cells(position, "C") = Worksheets(sheetNames(2)).Cells(requestsPointer, "K")
        Worksheets(1).Cells(position, "D") = Worksheets(sheetNames(2)).Cells(requestsPointer, "L")
        Worksheets(1).Cells(position, "E") = Worksheets(sheetNames(2)).Cells(requestsPointer, "M")
        Worksheets(1).Cells(position, "F") = Worksheets(sheetNames(2)).Cells(requestsPointer, "N")
        Worksheets(1).Cells(position, "G") = Worksheets(sheetNames(2)).Cells(requestsPointer, "B")
        Worksheets(1).Cells(position, "J") = Worksheets(sheetNames(2)).Cells(requestsPointer, "C")
        If Worksheets(sheetNames(2)).Cells(requestsPointer, "B").Value = "Product Loan" Or Worksheets(sheetNames(2)).Cells(requestsPointer, "B").Value = "Consulting" Then
            If Worksheets(sheetNames(2)).Cells(requestsPointer, "B").Value = "Product Loan" Then
                numberProduct = numberProduct + 1
            End If
            If Worksheets(sheetNames(2)).Cells(requestsPointer, "B").Value = "Consulting" Then
                numberConsult = numberConsult + 1
            End If
            Dim newRowPosition As Integer
            newRowPosition = Worksheets("Loans & Consult").Cells(Rows.Count, "A").End(xlUp).Row + 1
            Worksheets("Loans & Consult").Cells(newRowPosition, "A") = Worksheets("Loans & Consult").Cells(newRowPosition - 1, "A") + 1
            Worksheets("Loans & Consult").Cells(newRowPosition, "B") = Worksheets(sheetNames(2)).Cells(requestsPointer, "B")
            Worksheets("Loans & Consult").Cells(newRowPosition, "C") = Worksheets(sheetNames(2)).Cells(requestsPointer, "C")
            Worksheets("Loans & Consult").Cells(newRowPosition, "C").NumberFormat = "M/D/YYYY"
            Worksheets("Loans & Consult").Cells(newRowPosition, "G") = Worksheets(sheetNames(2)).Cells(requestsPointer, "G")
            Worksheets("Loans & Consult").Cells(newRowPosition, "H") = Worksheets(sheetNames(2)).Cells(requestsPointer, "H")
            Worksheets("Loans & Consult").Cells(newRowPosition, "I") = Worksheets(sheetNames(2)).Cells(requestsPointer, "I")
            Worksheets("Loans & Consult").Cells(newRowPosition, "J") = Worksheets(sheetNames(2)).Cells(requestsPointer, "J")
            Worksheets("Loans & Consult").Cells(newRowPosition, "K") = Worksheets(sheetNames(2)).Cells(requestsPointer, "K")
            Worksheets("Loans & Consult").Cells(newRowPosition, "L") = Worksheets(sheetNames(2)).Cells(requestsPointer, "L")
            Worksheets("Loans & Consult").Cells(newRowPosition, "M") = Worksheets(sheetNames(2)).Cells(requestsPointer, "M")
            Worksheets("Loans & Consult").Cells(newRowPosition, "N") = Worksheets(sheetNames(2)).Cells(requestsPointer, "N")

        End If
        If Worksheets(sheetNames(2)).Cells(requestsPointer, "B").Value = "Software" Then
            'Parse the Software Field with the Split Function
            Dim Softwares As Variant
            Softwares = Split(Worksheets(sheetNames(2)).Cells(requestsPointer, "G").Value, ",")
            Dim NumberOfSoftwares As Long
            NumberOfSoftwares = UBound(Softwares) - LBound(Softwares) + 1
            'Software Requested (From G to H)
            Worksheets(1).Cells(position, "H") = Worksheets(sheetNames(2)).Cells(requestsPointer, "G")
            'Number of Licenses (From H to I)
            Worksheets(1).Cells(position, "I") = Worksheets(sheetNames(2)).Cells(requestsPointer, "H")
            Dim softwareCounter As Integer

            For softwareCounter = 0 To NumberOfSoftwares - 1
                Dim currSoftware As String
                currSoftware = Softwares(softwareCounter)
                currSoftware = Trim(currSoftware)
                'L, M, and N in addition to S, T, and O populated with RW, EQ, and WQ
                If currSoftware = "Read & Write" Then
                    Worksheets(1).Cells(position, "L") = Worksheets(sheetNames(2)).Cells(requestsPointer, "H")
                    Worksheets(1).Cells(position, "S") = Worksheets(sheetNames(2)).Cells(requestsPointer, "H")

                End If
                If currSoftware = "EquatIO" Then
                    Worksheets(1).Cells(position, "M") = Worksheets(sheetNames(2)).Cells(requestsPointer, "H")
                    Worksheets(1).Cells(position, "T") = Worksheets(sheetNames(2)).Cells(requestsPointer, "H")
                End If

                If currSoftware = "WriQ" Then
                    Worksheets(1).Cells(position, "N") = Worksheets(sheetNames(2)).Cells(requestsPointer, "H")
                    Worksheets(1).Cells(position, "U") = Worksheets(sheetNames(2)).Cells(requestsPointer, "H")
                End If
            Next softwareCounter
            
            'Calculate Totals for the Software
            Worksheets(1).Cells(position, "O") = Worksheets(1).Cells(position, "L") + Worksheets(1).Cells(position, "M") + Worksheets(1).Cells(position, "N")
            Worksheets(1).Cells(position, "V") = Worksheets(1).Cells(position, "S") + Worksheets(1).Cells(position, "T") + Worksheets(1).Cells(position, "U")
            



        End If

    Next numberIndex
    
    Dim lastRow As Long
    lastRow = Worksheets(11).Cells(Rows.Count, "A").End(xlUp).Row
    Worksheets(11).Cells(lastRow, "B").Value = Worksheets(11).Cells(lastRow, "B").Value + numberProduct
    Worksheets(11).Cells(lastRow, "F").Value = Worksheets(11).Cells(lastRow, "F").Value + numberConsult
    
    Dim lastRowLoans As Long
    Dim lastRowConsults As Long
    lastRowLoans = Worksheets(12).Cells(Rows.Count, "A").End(xlUp).Row
    lastRowConsults = Worksheets(12).Cells(Rows.Count, "T").End(xlUp).Row
    Worksheets(12).Cells(lastRowLoans + 1, "A").Value = Left(myValue, 5) + " Totals"
    Worksheets(12).Cells(lastRowLoans + 1, "B").Value = numberProduct
    Worksheets(12).Cells(lastRowConsults + 1, "T").Value = Left(myValue, 5) + " Totals"
    Worksheets(12).Cells(lastRowConsults + 1, "U").Value = numberConsult

    Worksheets(12).Activate
    Worksheets(12).Range("A" + CStr(lastRowLoans + 1) + ":" + "B" + CStr(lastRowLoans + 1)).Select
    
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    Worksheets(12).Range("T" + CStr(lastRowConsults + 1) + ":" + "U" + CStr(lastRowConsults + 1)).Select
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    
    
    
    
    
    
    
    
    
    
End Sub
