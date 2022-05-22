' Originally from FindNewUsers
LastRowWeekly = Worksheets(sheetNames(8)).Cells(Rows.Count, "B").End(xlUp).Row
LastRowMonthly = Worksheets(sheetNames(9)).Cells(Rows.Count, "B").End(xlUp).Row


myValue = InputBox("What is today's date? (MM/DD/YYYY)")

Worksheets(sheetNames(9)).Cells(LastRowMonthly, "B") = Str(CLng(Worksheets(sheetNames(9)).Cells(LastRowMonthly, "B")) + counter)

Worksheets(sheetNames(8)).Cells(LastRowWeekly + 1, "B") = Worksheets(sheetNames(8)).Cells(LastRowWeekly, "B") + counter
Worksheets(sheetNames(8)).Cells(LastRowWeekly + 1, "A") = myValue
Worksheets(sheetNames(8)).Cells(LastRowWeekly + 1, "A").HorizontalAlignment = xlRight


