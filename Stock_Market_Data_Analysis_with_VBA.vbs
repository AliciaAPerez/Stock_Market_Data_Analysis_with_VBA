'Instructions
'Loop for each stock combining all info for each stock
'Column 1: The Ticker Symbol
'Column 2: Yearly Change
'Column 3: Percentage Change
'Column 4: Total Stock Volume
'Needs Conditional Formatting on Yearly Change
'Make Percentages look like percentages
'make it work for every sheet
'Bonus three items list with Ticker & Value
'#1: Greatest % Increase
'#2: Greatest % Decrease
'#3: Greatest Total Stock Volume

Sub StockMarket()

    
        Dim Ticker As String
        Dim Yearly As Double
        Dim Percentage As Double
        Dim Total As Double
        Dim ClosePrice As Double
        Dim OpenPrice As Double
        Dim i As Double
        Dim j As Double
        Dim Increase As Double
        Dim TickerIncrease As String
        Dim Decrease As Double
        Dim TickerDecrease As String
        Dim GreatestVol As Double
        Dim GreatestTicker As String
        
        
 For Each ws In Worksheets
        
        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Column Headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percentage Changed"
        ws.Range("L1").Value = "Total Stock Volume"
        
        'Set values to start
        'j is for the summary table entries
        j = 2
        OpenPrice = ws.Cells(2, 3).Value
        Increase = 0
        Decrease = 0
        GreatestVol = 0
        
        'Start Loop
        For i = 2 To LastRow
        
            'Find end of the group of stocks
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                'Pull Ticker info
                Ticker = ws.Cells(i, 1).Value
                'Find final price for stock
                ClosePrice = ws.Cells(i, 6).Value
                'Set yearly value
                Yearly = ClosePrice - OpenPrice
                
                    'Set percentage value
                    'cannot divide by zero, set up if statement for percentage
                    If OpenPrice <> 0 Then
                        Percentage = Yearly / OpenPrice
                        ElseIf OpenPrice = 0 Then
                            Percentage = 0
                    End If
                
                'Find Total of Stock volume
                Total = Total + ws.Cells(i, 7).Value
            
                'Set Cells for input
                ws.Cells(j, 9).Value = Ticker
                ws.Cells(j, 10).Value = Yearly
                ws.Cells(j, 11).Value = Percentage
                ws.Cells(j, 12).Value = Total
            
                    'For color of Yearly
                    If ws.Cells(j, 10).Value > 0 Then
                        ws.Cells(j, 10).Interior.ColorIndex = 4
                
                        ElseIf ws.Cells(j, 10).Value < 0 Then
                            ws.Cells(j, 10).Interior.ColorIndex = 3
                    
                    End If
                
                    'For Bonus greatest increase
                    If Percentage > Increase Then
                        Increase = Percentage
                        TickerIncrease = Ticker
                    End If
                
                    'For Bonus greatest decrease
                    If Percentage < Decrease Then
                        Decrease = Percentage
                        TickerDecrease = Ticker
                    End If
                
                    'For Bonus greatest total volume
                    If Total > GreatestVol Then
                        GreatestVol = Total
                        GreatestTicker = Ticker
                    End If
                                                
                'Reset all values
                Total = 0
                OpenPrice = ws.Cells(i + 1, 3).Value
                j = j + 1
                
            'refer back to initial if, we must toal out the sum once we're no longer in the same set
            Else
            Total = Total + ws.Cells(i, 7).Value
            
            'finally end the initial if statement for the group of stocks
            End If
            
        Next i
    
    'set formatting for percentages
    ws.Range("K:K").NumberFormat = "0.00%"
    
    'set up headers for bonus
    ws.Range("Q1").Value = "Value"
    ws.Range("P1").Value = "Ticker"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    'input values for bonus
    ws.Range("Q2").Value = Increase
    ws.Range("P2").Value = TickerIncrease
    ws.Range("Q3").Value = Decrease
    ws.Range("P3").Value = TickerDecrease
    ws.Range("Q4").Value = GreatestVol
    ws.Range("P4").Value = GreatestTicker
                   
    'set formatting for percentages
    ws.Range("Q2:Q3").NumberFormat = "0.00%"
    
    'set formatting for column size to make all headers legible
    ws.Range("J:L,O:O,Q:Q").Columns.AutoFit
    
    
   Next ws
                   
    
End Sub