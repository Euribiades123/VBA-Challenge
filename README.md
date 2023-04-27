# VBA-Challenge
Attribute VB_Name = "Module1"
'Create a script that loops through all the stocks for one year and outputs the following information:
'The ticker symbol
'Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
'The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
'The total stock volume of the stock.
'Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume".
'Make the appropriate adjustments to your VBA script to enable it to run on every worksheet (that is, every year) at once.
'Make sure to use conditional formatting that will highlight positive change in green and negative change in red.
'_______________________________________________________________________________________________________________________________________________
Sub VBA_Challenge()
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim Ticker As String
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalVolume As Double
    Dim SummaryRow As Double
    Dim TickerRow As Double
    
    'Loop through each worksheet
    For Each ws In Worksheets
        
        'Set initial values
        
        'Starting row for summary table
        SummaryRow = 2
        'Last row in worksheet
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Create summary table headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Volume"
        
        'Loop through each row in worksheet
        For i = 2 To LastRow
            
            'check if the ticker symbol changed then get Ticker symbol and Total Volume.
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Ticker = ws.Cells(i, 1).Value
                'totalVolume uses an excel function and calculates the total volume for each Ticker
                TotalVolume = Application.WorksheetFunction.SumIf(ws.Range("A:A"), Ticker, ws.Range("G:G"))
                
                'Get Open and Close Prices for each ticker
                TickerRow = Application.WorksheetFunction.Match(Ticker, ws.Range("A:A"), 0)
                OpenPrice = ws.Cells(TickerRow, 3).Value
                ClosePrice = ws.Cells(i, 6).Value
                
                'Calculate Yearly Change and Percent Change
                YearlyChange = ClosePrice - OpenPrice
                If OpenPrice <> 0 Then
                    PercentChange = YearlyChange / OpenPrice
                Else
                    PercentChange = 0
                End If
                
                'Add values to summary table
                ws.Cells(SummaryRow, 9).Value = Ticker
                ws.Cells(SummaryRow, 10).Value = YearlyChange
                ws.Cells(SummaryRow, 11).Value = PercentChange
                ws.Cells(SummaryRow, 12).Value = TotalVolume
                
                'Format PercentChange as percentage
                ws.Cells(SummaryRow, 11).NumberFormat = "0.00%"
                
                'Conditional formatting seting the colors on the yearly change column
                If YearlyChange > 0 Then
                    ws.Cells(SummaryRow, 10).Interior.ColorIndex = 4 'Green
                ElseIf YearlyChange < 0 Then
                    ws.Cells(SummaryRow, 10).Interior.ColorIndex = 3 'Red
                Else
                    ws.Cells(SummaryRow, 10).Interior.ColorIndex = 0 'No color
                End If
                
                'Conditional formatting seting the colors on the Percent Change column
                If PercentChange > 0 Then
                    ws.Cells(SummaryRow, 11).Interior.ColorIndex = 4 'Green
                ElseIf PercentChange < 0 Then
                    ws.Cells(SummaryRow, 11).Interior.ColorIndex = 3 'Red
                Else
                    ws.Cells(SummaryRow, 11).Interior.ColorIndex = 0 'No color
                End If
                
                'Move to next row in summary table
                SummaryRow = SummaryRow + 1
                
            End If
        Next i
        'End the first portion of the assignement
        '---------------------------------------------------------------------------
         'Add functionality to your script to return the stock with the
         ' "Greatest % increase", "Greatest % decrease",and "Greatest total volume".
        
        'Set Variable
        Dim GreatestIncrease As Double
        Dim GreatestDecrease As Double
        Dim GreatestVolume As Double
        
         'Create headers for greatest percent increase, Decrease and greatest volume
         ws.Range("N2").Value = "Greatest % Increase"
         ws.Range("N3").Value = "Greatest % Decrease"
         ws.Range("N4").Value = "Greatest Volume"
         ws.Range("O1").Value = "Ticker"
         ws.Range("P1").Value = "Value"
         
         'find the values for greatest increase, greatest decrease and largest volume
         GreatestIncrease = Application.WorksheetFunction.Max(ws.Range("K:K"))
         GreatestDecrease = Application.WorksheetFunction.Min(ws.Range("K:K"))
         GreatestVolume = Application.WorksheetFunction.Max(ws.Range("L:L"))
         
        'looping the summary data set
        For i = 2 To LastRow
        
        'Set values in table
         If ws.Cells(i, 11).Value = GreatestIncrease Then
            ws.Range("O2").Value = ws.Cells(i, 9).Value
            ws.Range("P2").Value = GreatestIncrease
            ws.Range("P2").NumberFormat = "0.00%"
         ElseIf ws.Cells(i, 11).Value = GreatestDecrease Then
            ws.Range("O3").Value = ws.Cells(i, 9).Value
            ws.Range("P3").Value = GreatestDecrease
            ws.Range("P3").NumberFormat = "0.00%"
        ElseIf ws.Cells(i, 12).Value = GreatestVolume Then
            ws.Range("O4").Value = ws.Cells(i, 9).Value
            ws.Range("P4").Value = GreatestVolume
                        
        End If
        Next i
                
    Next ws
    
End Sub





