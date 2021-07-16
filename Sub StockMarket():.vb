Sub StockMarket():

    ' Loop / Through All Worksheets
    For Each ws In Worksheets

        ' Column Headers / Data Field Labels
         ws.Cells(1,9).Value = "Ticker"
         ws.Cells(1,10).Value = "Yearly Change"
         ws.Cells(1,11).Value = "Percent Change"
         ws.Cells(1,12).Value = "Total Stock Volume"
         ws.Cells(1,16).Value = "Ticker"
         ws.Cells(1,17).Value = "Value"
         ws.Cells(2,15).Value = "Greatest % Increase"
         ws.Cells(3,15).Value = "Greatest % Decrease"
         ws.Cells(4,15).Value = "Greatest Total Volume"

        ' Set/Declare Initial Variables And Set Default/Baseline Variables
        Dim TickerSymbol As String
        Dim LastRow As Long
        Dim TotalStockVolume As Double
        TotalStockVolume = 0
        Dim OpeningPrice As Double
        Dim ClosingPrice As Double
        Dim PriceDifference As Double
        Dim PercentChange As Double
        Dim GreatestIncrease As Double
        GreatestIncrease = 0
        Dim GreatestDecrease As Double
        GreatestDecrease = 0
        Dim GreatestTotalVolume As Double
        GreatestTotalVolume = 0
        Dim Analysis As Long 
        Analysis = 2

        ' Define the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For allrows = 2 To LastRow

            ' Add To Ticker Total Volume
            TotalStockVolume = TotalStockVolume + ws.Cells(allrows, 7).Value
            ' Check if Same Ticker 
            If ws.Cells(allrows + 1, 1).Value <> ws.Cells(allrows, 1).Value Then

                ' Set Ticker Symbol
                TickerSymbol = ws.Cells(allrows, 1).Value
                ' Print The Ticker Symbol In The Summary Table
                ws.Cells(Analysis, 9).Value = ws.cells(allrows, 1).Value
                ' Print The Ticker Total Amount To The Summary Table
                ws.cells(Analysis, 12).Value = TotalStockVolume
                ' Reset Ticker Total
                TotalStockVolume = 0

                ' Set Opening Price , Opening Price and Yearly Change Name
                OpeningPrice = ws.Cells(2, 3).Value
                ClosingPrice = ws.Cells(allrows,6).Value
                PriceDifference = ClosingPrice - OpeningPrice
                ws.Cells(Analysis, 10).Value = PriceDifference

                ' Determine Percent Change
                If OpeningPrice = 0 Then
                    PercentChange = 0
                Else
                    PercentChange = PriceDifference / OpeningPrice
                End If
                ' Format To Include Percentage And Two Decimal Places
                ws.Cells(Analysis, 11).NumberFormat = "0.00%"
                ws.Cells(Analysis, 11).Value = PercentChange

                ' Conditional Formatting Highlight Positive (Green) / Negative (Red)
                If ws.Cells(Analysis, 10).Value >= 0 Then
                    ws.Cells(Analysis, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(Analysis, 10).Interior.ColorIndex = 3
                End If
            
                ' Add One To The Summary Table Row
                Analysis = Analysis + 1
                End If
            Next allrows

            ' Greatest % Increase, Greatest % Decrease and Greatest Total Volume
            'Define LastRow 
            LastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
        
            ' Start Loop For Final Results
            For Analysis = 2 To LastRow
                If ws.Cells(Analysis, 11).Value > ws.Cells(2, 17).Value Then
                    ws.Cells(2, 17).Value = ws.Cells(Analysis, 11).Value
                    ws.Cells(2, 16).Value = ws.Cells(Analysis, 9).Value
                End If

                If ws.Cells(Analysis, 11).Value < ws.Cells(3, 17).Value Then
                    ws.Cells(3, 17).Value = ws.Cells(Analysis, 11).Value
                    ws.Cells(3, 16).Value = ws.Cells(Analysis, 9).Value
                End If

                If ws.Cells(Analysis, 12).Value > ws.Cells(4, 17).Value Then
                    ws.Cells(4, 17).Value = ws.Cells(Analysis, 12).Value
                    ws.Cells(4, 16).Value = ws.Cells(Analysis, 9).Value
                End If

            Next Analysis
        ' Format Double To Include % Symbol And Two Decimal Places
            ws.Cells(2, 17).NumberFormat = "0.00%"
            ws.Cells(2, 17).NumberFormat = "0.00%"
            
    

    Next ws

End Sub