Sub stocks()

    For Each ws In Worksheets

        Dim Ticker As String
        Ticker = 0

        Dim ticker_type As Integer
        ticker_type = 2

        Dim open_amt As Double
        Dim close_amt As Double
        Dim yearly_change As Double
        Dim percentage_change As Double

        Dim tot_volume As Double
        tot_volume = 0

        open_amt = ws.Cells(2, 3).Value

        ws.Range("I1,P1") = "Ticker"
        ws.Range("J1") = "Yearly Change"
        ws.Range("K1") = "Percent Change"
        ws.Range("L1") = "Total Stock Volume"
        ws.Range("Q1") = "Value"
        ws.Range("O2") = "Greatest % Increase"
        ws.Range("O3") = "Greatest % Decrease"
        ws.Range("O4") = "Greatest Total Volume"

        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        For i = 2 To lastrow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                tot_volume = tot_volume + ws.Cells(i, 7).Value
                Ticker = ws.Cells(i, 1).Value
                ws.Range("I" & ticker_type).Value = Ticker
                close_amt = ws.Cells(i, 6).Value
                yearly_change = close_amt - open_amt
                ws.Range("J" & ticker_type).Value = yearly_change
                
                If open_amt = 0 Then
                    percentage_change = 0
                Else
                    percentage_change = yearly_change / open_amt
                End If
                ws.Range("K" & ticker_type).Value = percentage_change
                ws.Range("K" & ticker_type).NumberFormat = "0.00%"
                ws.Range("L" & ticker_type).Value = tot_volume
                tot_volume = 0
                
                open_amt = ws.Cells(i + 1, 3).Value
                ticker_type = ticker_type + 1
            Else
                tot_volume = tot_volume + ws.Cells(i, 7).Value
            End If
    Next i
            For i = 2 To lastrow
                If ws.Cells(i, 10) > 0 Then
                        ws.Cells(i, 10).Interior.Color = RGB(0, 255, 0)
                Else
                    ws.Cells(i, 10).Interior.Color = RGB(255, 0, 0)
                End If
            
        Next i
    
    Next ws

End Sub
