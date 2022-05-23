Sub copyRequestIntoCDR()
    Dim requestPointer As Variant
    For Each requestPointer In rowRequestNumbers
        Dim position As Integer
        position = 2
        Dim lastRowE As Integer
        lastRowE = Worksheets(1).Cells(Rows.Count, "E").End(xlUp).Row
        Dim District As String
        District = Worksheets(2).Cells(requestPointer, "M").Value
        While position <= lastRowE And StrComp(Worksheets(1).Cells(position, "E").Value, District) <> 1
            position = position + 1
        Wend
        Worksheets(1).Rows(position).Insert
        Worksheets(1).Cells(position, "G").ClearFormats
        Worksheets(1).Cells(position, "A") = "Portal"
        Worksheets(1).Cells(position, "B") = Worksheets(sheetNames(2)).Cells(requestPointer, "J")
        Worksheets(1).Cells(position, "C") = Worksheets(sheetNames(2)).Cells(requestPointer, "K")
        Worksheets(1).Cells(position, "D") = Worksheets(sheetNames(2)).Cells(requestPointer, "L")
        Worksheets(1).Cells(position, "E") = Worksheets(sheetNames(2)).Cells(requestPointer, "M")
        Worksheets(1).Cells(position, "F") = Worksheets(sheetNames(2)).Cells(requestPointer, "N")
        Worksheets(1).Cells(position, "G") = Worksheets(sheetNames(2)).Cells(requestPointer, "B")
        Worksheets(1).Cells(position, "J") = Worksheets(sheetNames(2)).Cells(requestPointer, "C")
        Worksheets(1).Cells(position, "J").NumberFormat = "M/D/YYYY"
        If Worksheets(sheetNames(2)).Cells(requestPointer, "B").Value = "Software" Then
            Dim Softwares As Variant
            Softwares = Split(Worksheets(sheetNames(2)).Cells(requestPointer, "G").Value, ",")
            'Software Requested (From G to H)
            Worksheets(1).Cells(position, "H") = Worksheets(sheetNames(2)).Cells(requestPointer, "G")
            'Number of Licenses (From H to I)
            Worksheets(1).Cells(position, "I") = Worksheets(sheetNames(2)).Cells(requestPointer, "H")
            Dim software As Variant
            For Each software In Softwares
                currSoftware = Trim(software)
          'L, M, and N in addition to S, T, and O populated with RW, EQ, and WQ
                If currSoftware = "Read & Write" Then
                    Worksheets(1).Cells(position, "L") = Worksheets(sheetNames(2)).Cells(requestPointer, "H")
                    Worksheets(1).Cells(position, "S") = Worksheets(sheetNames(2)).Cells(requestPointer, "H")

                End If
                If currSoftware = "EquatIO" Then
                    Worksheets(1).Cells(position, "M") = Worksheets(sheetNames(2)).Cells(requestPointer, "H")
                    Worksheets(1).Cells(position, "T") = Worksheets(sheetNames(2)).Cells(requestPointer, "H")
                End If

                If currSoftware = "WriQ" Then
                    Worksheets(1).Cells(position, "N") = Worksheets(sheetNames(2)).Cells(requestPointer, "H")
                    Worksheets(1).Cells(position, "U") = Worksheets(sheetNames(2)).Cells(requestPointer, "H")
                End If
            Next software
      
            Worksheets(1).Cells(position, "O") = Worksheets(1).Cells(position, "L") + Worksheets(1).Cells(position, "M") + Worksheets(1).Cells(position, "N")
            Worksheets(1).Cells(position, "V") = Worksheets(1).Cells(position, "S") + Worksheets(1).Cells(position, "T") + Worksheets(1).Cells(position, "U")
        End If
    Next requestPointer
    
    
    MsgBox "Done copying requests to CDR"


End Sub
