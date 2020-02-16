Sub VBA_challenge()


For Each ws In Worksheets


' Add titles to the first row of each column that we're about to fill.
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Value"


'Find and define the last used row in the sheet
Dim StockLastRow As Long
    
        StockLastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
'Find the last used row in the "<tickers>" and return it to the last row of the "Tickers" column
    For i = 2 To StockLastRow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            'Find and return the last row of the "Tickers" column
            Dim LastRowI As Long
            
            LastRowI = ws.Cells(Rows.Count, "I").End(xlUp).Row
            

            'Fill in the next row.
            ws.Cells(LastRowI + 1, 9).Value = ws.Cells(i, 1).Value
        End If
        Next i
        
'Count and add up the total in stock volume per Ticker per year

    'Create and define the variable used for the total of each volume
    Dim stockTotal As Double
    stockTotal = 0

    For i = 2 To StockLastRow
    'If they are the same ticker, then the value of the cell is added to the total.
        If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
            stockTotal = stockTotal + ws.Cells(i, 7).Value
            Else

                'Find and return the last row of the "Stock Total Volume" column
                Dim LastRowTotal As Long
                    
                    LastRowTotal = ws.Cells(Rows.Count, "L").End(xlUp).Row
                    

                'return the total to the last row of the "Stock Total Volume" column
                 ws.Cells(LastRowTotal + 1, 12).Value = stockTotal + ws.Cells(i, 7).Value

                'Reset the total for the next ticker
                stockTotal = 0

        End If

    Next i

'Find and return the yearly change and the percent change between the opening and closing prices of each year.

'Create variables for opening and closing prices
Dim opening As Double
Dim closing As Double
Dim YearlyChange As Double

    'For loop to search for opening price at the beginning of the year and closing price at the end of the year.
    For i = 2 To StockLastRow

        'Checks if Tickers are the same AND checks if opening has been assigned a price yet
        If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value And opening = 0 Then
        
        'assign value to opening price
        opening = ws.Cells(i, 3).Value
        
        'Checks if Ticker are different AND if closing has been assigned a price yet
        ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value And closing = 0 Then
            
            'Finda and return the last rows for both new columns.
            Dim LastRowJ As Long
            Dim LastRowK As Long

            
            LastRowJ = ws.Cells(Rows.Count, "J").End(xlUp).Row
            

            
            LastRowK = ws.Cells(Rows.Count, "K").End(xlUp).Row
            
            
            'Assign value to closing price
            closing = ws.Cells(i, 6).Value
            
            'Define YearlyChange as a variable
            YearlyChange = closing - opening
            
            
            'Assign YearlyChange to the last row of the "Yearly Change" column
            ws.Cells(LastRowJ + 1, 10).Value = YearlyChange
            

            

            If opening = 0 Then

                ws.Cells(LastRowK + 1, 11).Value = 0

                Else
                'Divide the YearlyChange by the opening price and then Format the value.
                ws.Cells(LastRowK + 1, 11).Value = FormatPercent(YearlyChange / opening)
                
            End If

            'Reset the opening and closing prices for the next Ticker
            opening = 0
            closing = 0

        End If
        'Create new conditional to find whether YearlyChange is greater or less than zero
        If YearlyChange > 0 Then

                'Format ws.Cells to green
                ws.Cells(LastRowJ + 1, 10).Interior.ColorIndex = 4
                
            ElseIf YearlyChange < 0 Then

                'Format ws.Cells to red
                ws.Cells(LastRowJ + 1, 10).Interior.ColorIndex = 3
        End If

    Next i
    
    Next ws
End Sub




