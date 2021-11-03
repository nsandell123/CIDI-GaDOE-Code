Option Explicit
Sub findRequestsDiff()
    Dim mainworkBook As Workbook
    'Used for Two Requests Sheets (Latest One and One Week Prior)
    Dim firstSheet As String
    Dim secondSheet As String
    Dim flag As Boolean
    'Used for For Loop
    Dim i As Integer
    'Used to find last row of column C of firstSheet and secondSheet respectively
    Dim LRowFirst As Long
    Dim LRowSecond As Long
    Set mainworkBook = ActiveWorkbook
    'Used to store row numbers of new requests in firstSheet
    Dim numbers() As Long
    ReDim numbers(5)
    'Used to index into counter array
    Dim counter As Long
    counter = 0
    'This first part is dedicated to finding the two Requests Sheets. The flag is used to distinguish between them.
    For i = 1 To mainworkBook.Sheets.Count
        If InStr(LCase(mainworkBook.Sheets(i).Name), "requests") <> 0 And Not (flag) Then
           firstSheet = mainworkBook.Sheets(i).Name
           flag = True

        End If
        If InStr(LCase(mainworkBook.Sheets(i).Name), "requests") <> 0 And (flag) Then
            secondSheet = mainworkBook.Sheets(i).Name

        End If
    Next i
    LRowFirst = Worksheets(firstSheet).Cells(Rows.Count, "C").End(xlUp).Row
    LRowSecond = Worksheets(secondSheet).Cells(Rows.Count, "C").End(xlUp).Row
    'This next section is dedicated to finding the requests diff in terms of the row number
    Dim Rng1 As Range
    Dim Rng2 As Range
    Set Rng1 = Worksheets(firstSheet).Range("C1:" + "C" + Right(Str(LRowFirst), Len(LRowFirst) - 1))
    Set Rng2 = Worksheets(secondSheet).Range("C1:" + "C" + Right(Str(LRowSecond), Len(LRowSecond) - 1))

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
    'MsgBox finalString

    'This next part will be dedicated to finding the location in alphabetical order
    Dim numberIndex As Integer
    Dim numberProduct As Integer
    Dim numberConsult As Integer
    For numberIndex = 0 To counter - 1
        Dim requestsPointer As Integer
        requestsPointer = numbers(numberIndex)
        Dim District As String
        District = Worksheets(firstSheet).Cells(requestsPointer, "M").Value
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
        Worksheets(1).Cells(position, "B") = Worksheets(firstSheet).Cells(requestsPointer, "J")
        Worksheets(1).Cells(position, "C") = Worksheets(firstSheet).Cells(requestsPointer, "K")
        Worksheets(1).Cells(position, "D") = Worksheets(firstSheet).Cells(requestsPointer, "L")
        Worksheets(1).Cells(position, "E") = Worksheets(firstSheet).Cells(requestsPointer, "M")
        Worksheets(1).Cells(position, "F") = Worksheets(firstSheet).Cells(requestsPointer, "N")
        Worksheets(1).Cells(position, "G") = Worksheets(firstSheet).Cells(requestsPointer, "B")
        Worksheets(1).Cells(position, "J") = Worksheets(firstSheet).Cells(requestsPointer, "C")
        If Worksheets(firstSheet).Cells(requestsPointer, "B").Value = "Product Loan" Or Worksheets(firstSheet).Cells(requestsPointer, "B").Value = "Consulting" Then
            If Worksheets(firstSheet).Cells(requestsPointer, "B").Value = "Product Loan" Then
                numberProduct = numberProduct + 1
            End If
            If Worksheets(firstSheet).Cells(requestsPointer, "B").Value = "Consulting" Then
                numberConsult = numberConsult + 1
            End If
            Dim newRowPosition As Integer
            newRowPosition = Worksheets("Loans & Consult").Cells(Rows.Count, "A").End(xlUp).Row + 1
            Worksheets("Loans & Consult").Cells(newRowPosition, "A") = Worksheets("Loans & Consult").Cells(newRowPosition - 1, "A")
            Worksheets("Loans & Consult").Cells(newRowPosition, "A") = Worksheets("Loans & Consult").Cells(newRowPosition, "A") + 1
            Worksheets("Loans & Consult").Cells(newRowPosition, "B") = Worksheets(firstSheet).Cells(requestsPointer, "B")
            Worksheets("Loans & Consult").Cells(newRowPosition, "C") = Worksheets(firstSheet).Cells(requestsPointer, "C")
            Worksheets("Loans & Consult").Cells(newRowPosition, "C").NumberFormat = "M/D/YYYY"
            Worksheets("Loans & Consult").Cells(newRowPosition, "G") = Worksheets(firstSheet).Cells(requestsPointer, "G")
            Worksheets("Loans & Consult").Cells(newRowPosition, "H") = Worksheets(firstSheet).Cells(requestsPointer, "H")
            Worksheets("Loans & Consult").Cells(newRowPosition, "I") = Worksheets(firstSheet).Cells(requestsPointer, "I")
            Worksheets("Loans & Consult").Cells(newRowPosition, "J") = Worksheets(firstSheet).Cells(requestsPointer, "J")
            Worksheets("Loans & Consult").Cells(newRowPosition, "K") = Worksheets(firstSheet).Cells(requestsPointer, "K")
            Worksheets("Loans & Consult").Cells(newRowPosition, "L") = Worksheets(firstSheet).Cells(requestsPointer, "L")
            Worksheets("Loans & Consult").Cells(newRowPosition, "M") = Worksheets(firstSheet).Cells(requestsPointer, "M")
            Worksheets("Loans & Consult").Cells(newRowPosition, "N") = Worksheets(firstSheet).Cells(requestsPointer, "N")
            
        End If
        If Worksheets(firstSheet).Cells(requestsPointer, "B").Value = "Software" Then
            'Parse the Software Field with the Split Function
            Dim Softwares As Variant
            Softwares = Split(Worksheets(firstSheet).Cells(requestsPointer, "G").Value, ",")
            Dim NumberOfSoftwares As Long
            NumberOfSoftwares = UBound(Softwares) - LBound(Softwares) + 1
            'Software Requested (From G to H)
            Worksheets(1).Cells(position, "H") = Worksheets(firstSheet).Cells(requestsPointer, "G")
            'Number of Licenses (From H to I)
            Worksheets(1).Cells(position, "I") = Worksheets(firstSheet).Cells(requestsPointer, "H")
            Dim softwareCounter As Integer

            For softwareCounter = 0 To NumberOfSoftwares - 1
                Dim currSoftware As String
                currSoftware = Softwares(softwareCounter)
                currSoftware = Trim(currSoftware)
                'L, M, and N in addition to S, T, and O populated with RW, EQ, and WQ
                If currSoftware = "Read & Write" Then
                    Worksheets(1).Cells(position, "L") = Worksheets(firstSheet).Cells(requestsPointer, "H")
                    Worksheets(1).Cells(position, "S") = Worksheets(firstSheet).Cells(requestsPointer, "H")
                    
                End If
                If currSoftware = "EquatIO" Then
                    Worksheets(1).Cells(position, "M") = Worksheets(firstSheet).Cells(requestsPointer, "H")
                    Worksheets(1).Cells(position, "T") = Worksheets(firstSheet).Cells(requestsPointer, "H")
                End If
                
                If currSoftware = "WriQ" Then
                    Worksheets(1).Cells(position, "N") = Worksheets(firstSheet).Cells(requestsPointer, "H")
                    Worksheets(1).Cells(position, "U") = Worksheets(firstSheet).Cells(requestsPointer, "H")
                End If
            Next softwareCounter
            
            
            
            
        
        End If
        
    Next numberIndex
    
    
    
    
    
    
    
End Sub