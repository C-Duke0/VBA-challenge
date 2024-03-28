Sub total()

    For Each ws In Worksheets

        LastRowI = ws.Cells(Rows.Count, 12).End(xlUp).Row
       
        Dim great_vol As Double
        Dim great_inc As Double
        Dim great_dec As Double

        great_vol = ws.Cells(2, 12).Value
        great_inc = ws.Cells(2, 11).Value
        great_dec = ws.Cells(2, 11).Value
            
            For i = 2 To LastRowI
            
                If ws.Cells(i, 12).Value > great_vol Then
                    great_vol = ws.Cells(i, 12).Value
                    ws.Range("P4").Value = ws.Cells(i, 9).Value
                Else
                    great_vol = great_vol
                End If

                If ws.Cells(i, 11).Value > great_inc Then
                    great_inc = ws.Cells(i, 11).Value
                    ws.Range("P2").Value = ws.Cells(i, 9).Value
                Else
                    great_inc = great_inc
                End If
                
                If ws.Cells(i, 11).Value < great_dec Then
                    great_dec = ws.Cells(i, 11).Value
                    ws.Range("P3").Value = ws.Cells(i, 9).Value
                Else
                    great_dec = great_dec
                End If
                
            ws.Range("Q4").Value = great_vol
            ws.Range("Q2").Value = great_inc
            ws.Range("Q3").Value = great_dec
            ws.Range("Q2", "Q3").NumberFormat = "0.00%"

            Next i
            
            
    Next ws
        
End Sub