Sub MinMax()

    Dim ws As Worksheet
    For Each ws In Worksheets

    ' Set variables to hold greatest % increase, % decrease, and total stock volume
    Dim dbGPI As Double
    Dim dbGPD As Double
    Dim dbGTSV As Double
    dbGPI = 1
    dbGPD = 1
    dbGTSV = 0


    ' Set variable to hold last row # and grab last row #
    Dim ngLastRow As Long
    ngLastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
    'Set up summary rows for greatest % increase, % decrease, and total stock volume
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Stock Volume"
    ws.Range("Q1").Value = "Value"
    

    ' Loop through all rows in summary table to find greatest % increase, put corresponding ticker symbol in 2nd summary table
    For i = 2 To ngLastRow
            If ws.Cells(i, 11).Value > dbGPI Then
            dbGPI = ws.Cells(i, 11).Value
            ws.Range("P2").Value = ws.Cells(i, 9).Value
    
            End If

            If ws.Cells(i, 11).Value < dbGPD Then
            dbGPD = ws.Cells(i, 11).Value
            ws.Range("P3").Value = ws.Cells(i, 9).Value
    
            End If

            If ws.Cells(i, 12).Value > dbGTSV Then
            dbGTSV = ws.Cells(i, 12).Value
            ws.Range("P4").Value = ws.Cells(i, 9).Value
    
            End If
    
    Next i
    
 ' Enter values in table
ws.Range("Q2").Value = FormatPercent(dbGPI)
ws.Range("Q3").Value = FormatPercent(dbGPD)
ws.Range("Q4").Value = dbGTSV
    
 
Next ws

End Sub
