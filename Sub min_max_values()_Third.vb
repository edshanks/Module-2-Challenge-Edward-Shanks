'run third
'this subroutine:
    'fills in max/min values and tickers in the third table
    'formats "Greatest % Increase" and "Greatest % Decrease" in the thrid table as percentages
    'formats "Greatest Total Volume" in the third table as a number with commas
    'widens columns to fit values and headers

Sub min_max_values()
    Dim lastrow As Long
    Dim ws As Worksheet
    
    For Each ws In Worksheets

        'counts number of rows in newly created table
        lastrow = ws.Cells(Rows.Count, 10).End(xlUp).row

        Dim min as Double
        Dim max As Double
        Dim total_volume as double

        'set initial max = to %change in first row
        max = ws.Cells(2, 11).Value

        'set initial min = to %change in first row
        min = ws.Cells(2, 11).Value
        
        'set initial greatest total volume to volume in first row
        total_volume = ws.Cells(2, 12).Value

         Dim ticker_max as string
         Dim ticker_min as string
         Dim ticker_maxVolume as string

        'loop to find max %change
        For row = 2 To lastrow
            
            If ws.Cells(row, 11).Value > max Then
                max = ws.Cells(row, 11).Value
                ticker_max = ws.Cells(row, 9).Value
                
            End If
        
        Next row

        'loop to find min %change
        For row = 2 To lastrow
            
            If ws.Cells(row, 11).Value < min Then
                min = ws.Cells(row, 11).Value
                ticker_min = ws.Cells(row, 9).Value
                
            End If
        
        Next row

        'loop to find greatest total volume
        For row = 2 To lastrow
            
            If ws.Cells(row, 12).Value > total_volume Then
                total_volume = ws.Cells(row, 12).Value
                ticker_maxVolume = ws.Cells(row, 9).Value
                
            End If
        
        Next row


        'adds min/max values to new table
        ws.Cells(2, 17).Value = max
        ws.Cells(3, 17).Value = min
        ws.Cells(4, 17).Value = total_volume

        'adds corresponding tickers to above values
        ws.Cells(2, 16).Value = ticker_max
        ws.Cells(3, 16).Value = ticker_min
        ws.Cells(4, 16).Value = ticker_maxVolume

        'widens all columns to fit
        ws.Columns("A:Q").AutoFit

    Next ws



End Sub
