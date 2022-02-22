'Note: Open VBA_Challenege.xlsm and run subscript in there
Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single
    yearValue = InputBox("What year would you like to run the analysis on?")
    startTime = Timer
    'Starting a timer after the user decides what year to run the analysis on'
    
    Dim tickers(12) As String
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
   'Initialized an array of all tickers
    Dim tickerIndex As Single
    tickerIndex = 0
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    Dim RowCount As Long
    Dim t As Byte
    Dim r As Long
    Dim dataRowStart As Byte
    dataRowStart = 4
    Dim dataRowEnd As Byte
    dataRowEnd = 15
    'Defined variables and arrays that will be used throughout the subscript

    Worksheets("All Stocks Analysis").Activate
    
    Cells(1, 1).Value = "All Stocks (" + yearValue + ")"
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    'Created header row

    Worksheets(yearValue).Activate
    
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    'Counted rows in desired dataset
    
    For t = 0 To 11
        tickerVolumes(t) = 0
    Next t
    'Looped through tickerVolumes to make sure variables in array are equal to 0

    For r = 2 To RowCount
        If Cells(r, 1).Value = tickers(tickerIndex) Then
        'Looping through all rows in sheet and checking to see if sheet ticker matches tickerIndex
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(r, 8).Value
            'Compiling the total volume for each ticker
            If Cells(r - 1, 1) <> tickers(tickerIndex) Then
                tickerStartingPrices(tickerIndex) = Cells(r, 6).Value
            End If
            'Getting the starting price for the year for the specific ticker
            If Cells(r + 1, 1) <> tickers(tickerIndex) Then
                tickerEndingPrices(tickerIndex) = Cells(r, 6).Value
                tickerIndex = tickerIndex + 1
            End If
            'Getting the ending price and switching tickerIndex if next ticker value does not match the current one
        End If
    Next r
    
    For t = 0 To 11
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + t, 1).Value = tickers(t)
        Cells(4 + t, 2).Value = tickerVolumes(t)
        Cells(4 + t, 3).Value = tickerEndingPrices(t) / tickerStartingPrices(t) - 1
        'Outputted collected data in analysis sheet
    Next t
    
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit
    'Formatted analysis sheet

    For i = dataRowStart To dataRowEnd
        If Cells(i, 3) > 0 Then
            Cells(i, 3).Interior.Color = vbGreen
        Else
            Cells(i, 3).Interior.Color = vbRed
        End If
    Next i
    'Created loop to color return column based on if stock was a gain or a loss
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
End Sub