Attribute VB_Name = "Module1"
Sub TickerOutput()

Dim ticker_name As String
Dim summary_ticker_row As Integer
Dim total_stock_volume As Variant

summary_ticker_row = 2
total_stock_volume = 0
Range("H1").Value = "Yearly_Open"
Range("I1").Value = "Yearly_Close"
Range("J1").Value = "Ticker"
Range("K1").Value = "Yearly_Change"
Range("L1").Value = "Percent_Change"
Range("M1").Value = "Total_Stock_Volume"



    For i = 2 To 22771
        'Find last row of new ticker
        If i = 2 Then
            Range("H" & 2).Value = Cells(i, 3).Value
        ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            'add ticker value to summary table
            ticker_name = Cells(i, 1).Value
            Range("J" & summary_ticker_row).Value = ticker_name
            'Find next ticker yearly open by checking row underneath
            Range("H" & summary_ticker_row + 1).Value = Cells(i + 1, 3).Value
            'Find yearly close
            Range("I" & summary_ticker_row).Value = Cells(i, 6).Value
            'add total stock volume to summary table
            Range("M" & summary_ticker_row).Value = total_stock_volume
            'Calculate yearly change
            Range("K" & summary_ticker_row).Value = Range("I" & summary_ticker_row).Value - Range("H" & summary_ticker_row).Value
            'Calculate percent change
            Range("L" & summary_ticker_row).Value = ((Range("I" & summary_ticker_row).Value - Range("H" & summary_ticker_row).Value) / Range("H" & summary_ticker_row).Value)
            'increment summary table row by 1
            summary_ticker_row = summary_ticker_row + 1
            'reset total stock volume for new ticker
            total_stock_volume = 0
        'Find first row of new ticker
        'ElseIf Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
            'yearly_open = Cells(i, 3).Value
            'Cells(yearly_open_row_increment, 8).Value = yearly_open
        Else
            'add total stock volume values for same tickers
            total_stock_volume = total_stock_volume + Cells(i, 7).Value
        End If
    Next i
    
    
End Sub
