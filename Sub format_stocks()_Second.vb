'run second
'this subroutine:
    'applies red/green conditional formatting for "Yearly Change Column"
    'formats "Percent Change" as percentages
    'formats "Total Stock Volume" as a number with commas
    'widens columns to fit data (not headers)
        'header fit formatting on next subroutine

Sub format_stocks()
Dim ws As Worksheet
Dim lastrow As Long

    For Each ws In Worksheets

        'format max/min % change as percentages
        ws.Range("Q2:Q3").NumberFormat = "0.00%"    

        'format Greatest Total Volume as # with commas
        ws.Range("Q4").NumberFormat = "#,###"

        'center all column headers
        ws.Range("A1:Q1").HorizontalAlignment = xlCenter
        
        'counts rows in stock summary table
        lastrow = ws.Cells(Rows.Count, 9).End(xlUp).row
        
        'loops through newly created tables on each sheet
        For row = 2 To lastrow
        
            'formate Percent Change as percentages
            ws.Cells(row, 11).NumberFormat = "0.00%"

            'adds commas to total stock volume
            ws.Cells(row, 12).NumberFormat = "#,###"

            'red/green conditional formatting for yearly change
            If ws.Cells(row, 10).Value > 0 Then
                ws.Cells(row, 10).Interior.ColorIndex = 4
            
            ElseIf ws.Cells(row, 10).Value < 0 Then
                ws.Cells(row, 10).Interior.ColorIndex = 3
                
            End If

        
        
        Next row


        'widens all columns to fit
        'ws.Columns("A:Q").AutoFit




    Next ws




End Sub
