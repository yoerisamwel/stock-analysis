Attribute VB_Name = "Module1"
Sub stock_analysis()

Dim ws As Worksheet

    For Each ws In Worksheets
        
        Dim WorksheetName As String
        Dim ticker_symbol As String
        Dim ticker_symbol_vol_t As Double
        Dim ticker_symbol_change_t As Double
        Dim ticker_symbol_change_t_old As Double
        Dim ticker_symbol_change_t_new As Double
        Dim ticker_symbol_change_t_perc As Double
        Dim Summery_Table_Row As Integer
        Dim Max As Double
        Dim Maxper As Double
        Dim Minper As Double
        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        ws.Range("I1").EntireColumn.Insert
        ws.Range("I1").EntireColumn.Insert
        ws.Range("I1").EntireColumn.Insert
        ws.Range("I1").EntireColumn.Insert
        ws.Cells(1, 11).Value = "total_stock_volume"
        ws.Cells(1, 10).Value = "ticker_symbol"
        ws.Cells(1, 13).Value = "yearly_change"
        ws.Cells(1, 12).Value = "percent_change"
        ticker_symbol_vol_t = 0
        ticker_symbol_change_t = 0
        ticker_symbol_change_t_old = 0
        ticker_symbol_change_t_new = 0
        ticker_symbol_change_t_perc = 0
        
        Summery_Table_Row = 2
    
        For I = 2 To LastRow
     
            If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
                ticker_symbol = Cells(I, 1).Value
                ticker_symbol_vol_t = ticker_symbol_vol_t + ws.Cells(I, 7).Value
                ticker_symbol_change_t_old = ticker_symbol_change_t_old + ws.Cells(I, 3).Value
                ticker_symbol_change_t_new = ticker_symbol_change_t_new + ws.Cells(I, 6).Value
                ticker_symbol_change_t = ticker_symbol_change_t_old - ticker_symbol_change_t_new
                If ticker_symbol_change_t_old > 0 Then
                    ticker_symbol_change_t_perc = (ticker_symbol_change_t_old - ticker_symbol_change_t_new) / ticker_symbol_change_t_old * 100
                Else
                    ticker_symbol_change_t_perc = (0.01 - ticker_symbol_change_t_new) / 0.01 * 100
                End If
                ws.Range("J" & Summery_Table_Row).Value = ticker_symbol
                ws.Range("K" & Summery_Table_Row).Value = ticker_symbol_vol_t
                ws.Range("M" & Summery_Table_Row).Value = ticker_symbol_change_t
                ws.Range("L" & Summery_Table_Row).Value = ticker_symbol_change_t_perc
                
                Summery_Table_Row = Summery_Table_Row + 1
                ticker_symbol_vol_t = 0
                ticker_symbol_change_t = 0
            Else
                ticker_symbol_vol_total = ticker_symbol_vol_total + ws.Cells(I, 7).Value
            End If
        
        Next I
        ws.Range("L2:L" & LastRow).NumberFormat = "0.00%"
        
        LastRow2 = ws.Cells(Rows.Count, 10).End(xlUp).Row
        
        For x = 2 To LastRow2
            If ws.Cells(x, 12) <= 0 Then
                ws.Cells(x, 12).Interior.ColorIndex = 3
            ElseIf ws.Cells(x, 12) = 0 Then
                ws.Cells(x, 12).Interior.ColorIndex = 2
            ElseIf ws.Cells(x, 12) > 0 Then
                ws.Cells(x, 12).Interior.ColorIndex = 4
            End If
        
        Next x
        
        For y = 2 To LastRow2
            If ws.Cells(y, 13) <= 0 Then
                ws.Cells(y, 13).Interior.ColorIndex = 3
            ElseIf ws.Cells(y, 13) = 0 Then
                ws.Cells(y, 13).Interior.ColorIndex = 2
            ElseIf ws.Cells(y, 13) > 0 Then
                ws.Cells(y, 13).Interior.ColorIndex = 4
            End If
        
        Next y
        
        Max = 0
        Maxper = 0
        Minper = 0
        
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        For max_for = 2 To LastRow2
            If ws.Cells(max_for, 13).Value >= Max Then
                Max = ws.Cells(max_for, 13).Value
                ws.Range("P4").Value = Max
            End If
        Next max_for
        
        For max_per = 2 To LastRow2
            If ws.Cells(max_per, 12).Value >= Maxper Then
                Maxper = ws.Cells(max_per, 12).Value
                ws.Range("P2").Value = Maxper
            End If
        Next max_per
        
        For min_per = 2 To LastRow2
            If ws.Cells(min_per, 12).Value <= Minper Then
                Minper = ws.Cells(min_per, 12).Value
                ws.Range("P3").Value = Minper
            End If
        Next min_per
        
        ws.Range("P2").NumberFormat = "0.00%"
        ws.Range("P3").NumberFormat = "0.00%"
    
    Next ws

End Sub
