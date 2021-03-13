Sub stockData()
    'Find the last row of the data set
    Dim lastRow As Long
    'Current ticker
    Dim current As String
    'Yearly opening price of current ticker
    Dim opening As Double
    'Yearly closing price of current ticker
    Dim closing As Double
    'Total volume of current ticker
    Dim totalVolume As Double
    'Percent change between open and close price
    Dim percentChange As Double
    'Keep track of row for insert to new table
    Dim RowIndex As Integer
    'Counter in case opening price is 0
    Dim j As Long
    'Variables for keeping track of greatest increase/decrease/volume
    Dim incTicker As String
    Dim incValue As Double
    Dim decTicker As String
    Dim decValue As Double
    Dim volTicker As String
    Dim volValue As Double
    
    
    For Each ws In Worksheets
        'find the last row in the current worksheet
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'Reset RowIndex for current worksheet
        RowIndex = 2
        'Reset greatest increase, decrease, volume trackers
        incValue = 0
        decValue = 0
        volValue = 0
        
        'Column Headers
        ws.Range("J1") = "Ticker"
        ws.Range("K1") = "Yearly Change"
        ws.Range("L1") = "Percent Change"
        ws.Range("M1") = "Total Stock Volume"
        ws.Range("O2") = "Greatest % Increase"
        ws.Range("O3") = "Greatest % Decrease"
        ws.Range("O4") = "Greatest Total Volume"
        ws.Range("P1") = "Ticker"
        ws.Range("Q1") = "Value"
        
        'Loop through every row in the sheet
        For i = 2 To lastRow
            'Add current row volume to the total
            totalVolume = totalVolume + ws.Cells(i, 7)
            'If the row is the first of its ticker value
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                'Get the current ticker
                current = ws.Cells(i, 1).Value
                'Get the open price
                opening = ws.Cells(i, 3).Value
                j = i
                While ws.Cells(j, 3).Value <= 0 And ws.Cells(j, 1).Value = current
                    opening = ws.Cells(j, 3).Value
                    j = j + 1
                Wend
            ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                'Get the closing price
                closing = ws.Cells(i, 6).Value
                
                'Write current ticker into new table
                ws.Cells(RowIndex, 10).Value = current
                'Calculate yearly change in price
                ws.Cells(RowIndex, 11).Value = closing - opening
                'Calculate percent change from year open to year close
                If opening > 0 Then
                    ws.Cells(RowIndex, 12).Value = (closing - opening) / opening
                Else
                    ws.Cells(RowIndex, 12).Value = 0
                End If
                'Format the percentage cell with a fill color for + or -
                If ws.Cells(RowIndex, 12).Value < 0 Then
                    ws.Cells(RowIndex, 12).Interior.ColorIndex = 3
                ElseIf ws.Cells(RowIndex, 12).Value > 0 Then
                    ws.Cells(RowIndex, 12).Interior.ColorIndex = 4
                End If
                'Format the column showing a percentage correctly
                ws.Cells(RowIndex, 12).NumberFormat = "0.00%"
                ws.Cells(RowIndex, 13).Value = totalVolume
                
                'Check if the ticker that was just written is the
                'greatest increase, decrease, or volume
                If ws.Cells(RowIndex, 12).Value > incValue Then
                    incValue = ws.Cells(RowIndex, 12).Value
                    incTicker = current
                End If
                If ws.Cells(RowIndex, 12).Value < decValue Then
                    decValue = ws.Cells(RowIndex, 12).Value
                    decTicker = current
                End If
                If totalVolume > volValue Then
                    volValue = totalVolume
                    volTicker = current
                End If
                
                'Reset total to zero for next ticker
                totalVolume = 0
                'Move to next row index
                RowIndex = RowIndex + 1
            End If
        Next i
        'Fill greatest increase, decrease volume table
        ws.Range("P2").Value = incTicker
        ws.Range("Q2").Value = incValue
        ws.Range("P3").Value = decTicker
        ws.Range("Q3").Value = decValue
        ws.Range("P4").Value = volTicker
        ws.Range("Q4").Value = volValue
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        'AutoFit all columns written to
        ws.Columns("J:Q").AutoFit
    Next ws

End Sub
