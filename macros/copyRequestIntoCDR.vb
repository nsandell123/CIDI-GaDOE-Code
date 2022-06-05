Sub copyRequestIntoCDR()
    Dim requestPointer As Variant
    If numberNewRequests = 0 Then
        MsgBox "No New Requests"
        Exit Sub
    End If
    For Each requestPointer In rowRequestNumbers
        Dim position As Integer
        position = 2
        Dim lastRowE As Integer
        lastRowE = Worksheets(2).Cells(Rows.Count, "E").End(xlUp).Row
        Dim District As String
        District = Worksheets(3).Cells(requestPointer, "M").Value
        While position <= lastRowE And StrComp(Worksheets(2).Cells(position, "E").Value, District) <> 1
            position = position + 1
        Wend
        Worksheets(2).Rows(position).Insert
        Worksheets(2).Cells(position, "G").ClearFormats
        Worksheets(2).Cells(position, "A") = "Portal"
        Worksheets(2).Cells(position, "B") = Worksheets(3).Cells(requestPointer, "J")
        Worksheets(2).Cells(position, "C") = Worksheets(3).Cells(requestPointer, "K")
        Worksheets(2).Cells(position, "D") = Worksheets(3).Cells(requestPointer, "L")
        Worksheets(2).Cells(position, "E") = Worksheets(3).Cells(requestPointer, "M")
        Worksheets(2).Cells(position, "F") = Worksheets(3).Cells(requestPointer, "N")
        Worksheets(2).Cells(position, "G") = Worksheets(3).Cells(requestPointer, "B")
        Worksheets(2).Cells(position, "J") = Worksheets(3).Cells(requestPointer, "C")
        Worksheets(2).Cells(position, "J").NumberFormat = "M/D/YYYY"
        If Worksheets(3).Cells(requestPointer, "B").Value = "Product Loan" Or Worksheets(3).Cells(requestPointer, "B").Value = "Consulting" Then
            Dim newRowPosition As Integer
            newRowPosition = Worksheets("Loans & Consults").Cells(Rows.Count, "A").End(xlUp).Row + 1
            Worksheets("Loans & Consults").Cells(newRowPosition, "A") = Worksheets("Loans & Consults").Cells(newRowPosition - 1, "A") + 1
            Worksheets("Loans & Consults").Cells(newRowPosition, "B") = Worksheets(3).Cells(requestPointer, "B")
            Worksheets("Loans & Consults").Cells(newRowPosition, "C") = Worksheets(3).Cells(requestPointer, "C")
            Worksheets("Loans & Consults").Cells(newRowPosition, "C").NumberFormat = "M/D/YYYY"
            Worksheets("Loans & Consults").Cells(newRowPosition, "G") = Worksheets(3).Cells(requestPointer, "G")
            Worksheets("Loans & Consults").Cells(newRowPosition, "H") = Worksheets(3).Cells(requestPointer, "H")
            Worksheets("Loans & Consults").Cells(newRowPosition, "I") = Worksheets(3).Cells(requestPointer, "I")
            Worksheets("Loans & Consults").Cells(newRowPosition, "J") = Worksheets(3).Cells(requestPointer, "J")
            Worksheets("Loans & Consults").Cells(newRowPosition, "K") = Worksheets(3).Cells(requestPointer, "K")
            Worksheets("Loans & Consults").Cells(newRowPosition, "L") = Worksheets(3).Cells(requestPointer, "L")
            Worksheets("Loans & Consults").Cells(newRowPosition, "M") = Worksheets(3).Cells(requestPointer, "M")
            Worksheets("Loans & Consults").Cells(newRowPosition, "N") = Worksheets(3).Cells(requestPointer, "N")
            Worksheets("Loans & Consults").Range(Worksheets("Loans & Consults").Cells(newRowPosition, "A"), Worksheets("Loans & Consults").Cells(newRowPosition, "A").Offset(0, 13)).Borders.LineStyle = xlContinuous
            Worksheets("Loans & Consults").Range(Worksheets("Loans & Consults").Cells(newRowPosition, "A"), Worksheets("Loans & Consults").Cells(newRowPosition, "A").Offset(0, 13)).Borders.Weight = xlThin
        End If
        If Worksheets(3).Cells(requestPointer, "B").Value = "Software" Then
            Dim Softwares As Variant
            Softwares = Split(Worksheets(3).Cells(requestPointer, "G").Value, ",")
            'Software Requested (From G to H)
            Worksheets(2).Cells(position, "H") = Worksheets(3).Cells(requestPointer, "G")
            'Number of Licenses (From H to I)
            Worksheets(2).Cells(position, "I") = Worksheets(3).Cells(requestPointer, "H")
            Dim software As Variant
            For Each software In Softwares
                currSoftware = Trim(software)
                'L, M, and N in addition to S, T, and O populated with RW, EQ, and WQ
                If currSoftware = "Read & Write" Then
                    Worksheets(2).Cells(position, "L") = Worksheets(3).Cells(requestPointer, "H")
                    Worksheets(2).Cells(position, "S") = Worksheets(3).Cells(requestPointer, "H")

                End If
                If currSoftware = "EquatIO" Then
                    Worksheets(2).Cells(position, "M") = Worksheets(3).Cells(requestPointer, "H")
                    Worksheets(2).Cells(position, "T") = Worksheets(3).Cells(requestPointer, "H")
                End If

                If currSoftware = "WriQ" Then
                    Worksheets(2).Cells(position, "N") = Worksheets(3).Cells(requestPointer, "H")
                    Worksheets(2).Cells(position, "U") = Worksheets(3).Cells(requestPointer, "H")
                End If
            Next software

            Worksheets(2).Cells(position, "O") = Worksheets(2).Cells(position, "L") + Worksheets(2).Cells(position, "M") + Worksheets(2).Cells(position, "N")
            Worksheets(2).Cells(position, "V") = Worksheets(2).Cells(position, "S") + Worksheets(2).Cells(position, "T") + Worksheets(2).Cells(position, "U")
        End If
    Next requestPointer
    
    
    MsgBox "Done copying requests to CDR and Loans & Consults"


End Sub
