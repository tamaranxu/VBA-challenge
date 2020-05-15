Sub Stocks()

    'run script on every sheet
    For Each ws In Worksheets
        'add column labels to right of main table (ticker, yearly change, percent change, total stock volume, opening and closing)
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
                
        'keep track of the location for each ticker name in the summary table
        Dim Summary_Table_Row As Double
        'set starting row (initial value)
        Summary_Table_Row = 2
        'create variable to hold ticker name
        Dim Ticker As String
        Dim Opening_Value As Double
        Dim Closing_Value As Double
        'assign variable for stock count
        Dim Stock_Volume As Double
                    
        'use last row format so don't have to count/scroll to bottom
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
            'loop through all rows, excluding header row, checking for value in column 1
            For i = 2 To LastRow
                'assign value for ticker
                Ticker = ws.Cells(i, 1).Value
                'create new check for value of open on first row of each ticker
                If ws.Cells(i - 1, 1).Value <> Ticker Then
                    Opening_Value = ws.Cells(i, 3).Value
                    'set value for the stock count
                    Stock_Volume = ws.Cells(i, 7).Value
                'check for each row of ticker
                ElseIf ws.Cells(i - 1, 1).Value = Ticker And ws.Cells(i + 1, 1).Value = Ticker Then
                    'add value in G to Stock_Volume
                    Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
                'check if we are still within the same ticker name, if we are not...
                ElseIf ws.Cells(i + 1, 1).Value <> Ticker Then
                    'print ticker name in column I
                    ws.Range("I" & Summary_Table_Row).Value = Ticker
                    Closing_Value = ws.Cells(i, 6).Value
                    'print (closing - opening) in column j
                    ws.Range("J" & Summary_Table_Row).Value = Closing_Value - Opening_Value
                        'conditional formatting for J (positive in green; negative in red)
                        If ws.Range("J" & Summary_Table_Row).Value > 0 Then
                            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                        ElseIf ws.Range("J" & Summary_Table_Row).Value < 0 Then
                            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                        End If
                    'format K to have 2 decimal places and percentage
                    ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                    'print percent change in column K
                        If Opening_Value > 0 Then
                            ws.Range("K" & Summary_Table_Row).Value = (ws.Range("J" & Summary_Table_Row).Value) / Opening_Value
                        Else
                            'condition if value in K is 0
                            ws.Range("K" & Summary_Table_Row).Value = 0
                        End If
                    'add value in G to Stock_Volume
                    Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
                    'print Stock_Volume to L
                    ws.Range("L" & Summary_Table_Row).Value = Stock_Volume
                    'add one to the summary table row
                    Summary_Table_Row = Summary_Table_Row + 1
                    're-set variables to 0
                    Stock_Volume = 0
                    Opening_Value = 0
                    Closing_Value = 0
                End If
            Next i
        
        'set new last row variable for K & L rows
        LastRowTotal = ws.Cells(Rows.Count, 9).End(xlUp).Row
        'set up max and min variables for K/I, start at 0
        Dim MaxChange As Double
        MaxChange = 0
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        Dim MinChange As Double
        MinChange = 0
        'set up max variable for L/I
        Dim MaxVolume As Double
        MaxVolume = 0
            'new for loop to look through total rows (I to L); something to denote last row of data in those columns (exit or else)
            For i = 2 To LastRowTotal
                'if K(11) > max, set that value to max and P(16) to I(17)
                If ws.Cells(i, 11).Value >= 0 And ws.Cells(i, 11).Value >= MaxChange Then
                    ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                    ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
                    MaxChange = ws.Cells(i, 11).Value
                'if K < min, set that value to min and P to I
                ElseIf ws.Cells(i, 11).Value <= 0 And ws.Cells(i, 11).Value <= MinChange Then
                    ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                    ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
                    MinChange = ws.Cells(i, 11).Value
                End If
                
                'if L(12) > maxVolume, set that value to maxL and P(16) to I(17)
                If ws.Cells(i, 12).Value > MaxVolume Then
                    ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                    ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
                    MaxVolume = ws.Cells(i, 12).Value
                End If
            Next i
        ws.Columns("A:Q").AutoFit
    Next ws
End Sub
