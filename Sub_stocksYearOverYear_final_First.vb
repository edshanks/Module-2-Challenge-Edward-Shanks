'run first
'this subroutine: 
    'generates the summary table for all tickers on a sheet (table 2)
    'generates the empty max/min table (table 3)

Sub stocksYearOverYear()
    Dim ticker As String
    Dim open_price As Double
    Dim close_price As Double
    Dim total_volume As Double
    Dim daily_volume As Double
    Dim row As Long
    
    Dim open_count As Integer
    

    Dim lastrow As Long
    Dim ws As Worksheet

    For Each ws In Worksheets
        'create headers in new table each sheet
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"

        'adds row headers "Greatest % Increase", "Greatest % Decrease", "Greatest Total Volume" to max/min table
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        'adds column headers "Ticker" and "Value" to max/min table
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"


        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).row 'number of rows'
        
        open_count = 0 'counts # of times open price recorded in new column
                       'used to position data entry in new table
        
        total_volume = 0
        
        For row = 2 To lastrow
            ticker = ws.Cells(row, 1)
            daily_volume = ws.Cells(row, 7).Value
            total_volume = total_volume + daily_volume
        

            'determines first day of year and pulls year open price
            If ticker <> ws.Cells(row - 1, 1) Then



                'uses count to determine row position in new column
                open_count = open_count + 1
                
                'reads in ticker for new column
                ws.Cells(open_count + 1, 9) = ticker

                'reads in year open from data
                open_price = ws.Cells(row, 3)

                
            'determines last day of year and pulls year close price
            ElseIf ticker <> ws.Cells(row + 1, 1) Then
                
                'reads in year close from data
                close_price = ws.Cells(row, 6)

                'puts yearly change in new column''''''''''''''''''''''''
                ws.Cells(open_count + 1, 10) = close_price - open_price

                'puts percent change in new column
                ws.Cells(open_count + 1, 11) = (close_price - open_price) / open_price

                'puts total volume in new column
                ws.Cells(open_count + 1, 12) = total_volume
                
                'resets total volume for next ticker
                total_volume = 0
                

            End If
            

        Next row

    Next ws

    

    
End Sub
