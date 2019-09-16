Sub stocktest()


For Each ws In Worksheets

    ' Set variable to hold last row # and grab last row #
    Dim ngLastRow As Long
    ngLastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ' Set column location each stock's information
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

    ' Set first row of stock information, set as integer so rows can be added to it
    Dim Summary_Row As Integer
    Summary_Row = 2

    ' Set variables to hold ticker symbol, open/close values, percentage change, and total stock volume
    Dim strTickSym As String
    Dim dbOpen As Double
    Dim dbClose As Double
    Dim dbYearlychange As Double
    Dim varPercentagechange As Variant
    Dim dbTotalstockvolume As Double
    dbTotalstockvolume = 0

    ' Set variable to hold first row of range for each stock
    Dim iStartrow As Double
    iStartrow = 2

    ' Loop through all ticker symbols
    For i = 2 To ngLastRow
    ' Set ticker symbol
    strTickSym = ws.Cells(i, 1).Value
            
        ' Check whether ticker symbo matches the one below it, if it doesn't:
        If ws.Cells(i + 1, 1).Value <> strTickSym Then
            
            ' Set open & close dates
            dbOpen = ws.Cells(iStartrow, 3).Value
            dbClose = ws.Cells(i, 6).Value
                    
            ' Calculate yearly change
            dbYearlychange = (dbClose - dbOpen)
                    
            ' Calculate percentage change
            varPercentagechange = 0
            If dbOpen <> 0 Then
            varPercentagechange = (dbClose / dbOpen) - 1
            End If
            
            ' Add to total stock volume
            dbTotalstockvolume = ws.Cells(i, 7).Value + dbTotalstockvolume
                    
            ' Print ticker symbol to summary row
            ws.Cells(Summary_Row, 9).Value = strTickSym
                    
            ' Print yearly change to summary row
            ws.Cells(Summary_Row, 10).Value = dbYearlychange
                    
            ' Format color of cells to reflect +/- changes
                If dbYearlychange < 0 Then
                        
                    ws.Cells(Summary_Row, 10).Interior.ColorIndex = 3
                    Else: ws.Cells(Summary_Row, 10).Interior.ColorIndex = 4
                        
                End If
                
            ' Print percentage change to summary row and format as percent
            ws.Cells(Summary_Row, 11).Value = FormatPercent(varPercentagechange)
            
            ' Print total stock volume to summary row
            ws.Cells(Summary_Row, 12).Value = dbTotalstockvolume
         
            ' Add one row to summary columns
            Summary_Row = Summary_Row + 1
            
            ' Reset starting row
            iStartrow = i + 1
                    
            ' Reset ticker symbol
            strTickSym = ws.Cells(i + 1, 1).Value
                    
            ' Reset total stock volumen to 0
            dbTotalstockvolume = 0
        
        End If
    
    Next i

Next ws

End Sub
