Sub StockMarket()
    
    'Loop through all worksheets
    For Each ws In Worksheets
    
        'created variables to hold ticker symbol, total stock volume,the start of summary_row, totalvolume, opening and closing price, and year and percent change
        Dim ticker As String
        Dim last_row As Long
        Dim summary_row As Integer
        Dim TotalVolume As Double
        Dim openingprice As Double
        Dim closingprice As Double
        Dim yearchange As Double
        Dim percentchange As Double
    
        'determines the last row of the spreadsheet
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'initialize the starting point for the summary row, volume, percent change, opening price
        summary_row = 2
        TotalVolume = 0
        percentchange = 0
        openingprice = Cells(2, 3).Value
    
        'create summary table headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        'loops through the spreadsheet
        For i = 2 To last_row
    
            'add stock volume to total_volume
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value
    
            'check to see if ticker below the current cell was the same, if not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                'get the current row's ticker symbol
                ticker = ws.Cells(i, 1).Value
                'get the current row's closing price
                closingprice = ws.Cells(i, 6).Value
                yearchange = closingprice - openingprice
    
                'put ticker symbol, total volume, yearchange into summary table
                ws.Range("I" & summary_row).Value = ticker
                ws.Range("L" & summary_row).Value = TotalVolume
                ws.Range("J" & summary_row).Value = yearchange
                
                'change color based on negative or positive change
                If yearchange >= 0 Then
                    ws.Range("J" & summary_row).Interior.ColorIndex = 4
    
                Else
                    ws.Range("J" & summary_row).Interior.ColorIndex = 3
                    
                End If
                
                'calculate percent change
                If openingprice > 0 Then
                    percentchange = yearchange / openingprice
                    
                    'put percentchange in summary table
                    ws.Range("K" & summary_row).Value = percentchange
                    
                    'change format to percent
                    ws.Range("K" & summary_row).NumberFormat = "0.00%"
                
                Else
                    percentchange = NA
                    'next row is the opening price for the next opening
                    openingprice = ws.Cells(i + 1, 3)
    
                End If
                
                'added 1 to the summary row
                summary_row = summary_row + 1
                
                'redefine openingprice
                openingprice = Cells(i + 1, 3).Value
    
                'zero out total_volume every time ticker changes
                TotalVolume = 0
    
            End If
            
                
        Next i

    Next ws
    
End Sub






