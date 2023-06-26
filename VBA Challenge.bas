Attribute VB_Name = "Module1"
Sub stocks():

    Const FIRST_DATA_ROW As Integer = 2
    Const INPUT_TICKER_COL As Integer = 1
    Const OUTPUT_TICKER_COL As Integer = 9
    Const OUTPUT_CHANGE_COL As Integer = 10
    Const OUTPUT_PERCENT_COL As Integer = 11
    Const OUTPUT_TOTAL_COL As Integer = 12
    
    
    
    

    ' dim variables
    Dim current_ticker As String
    Dim next_ticker As String
    Dim last_row As Double
    Dim input_row As Double
    Dim output_row As Double
    Dim yearly_change As Double
    Dim closing_price As Double
    Dim opening_price As Double
    Dim percent_change As Double
    Dim volume As Double
    Dim high As Double
    Dim greatest_decrease As Double
    Dim total As Double
    Dim high_ticker As String
    Dim low_ticker As String
    Dim total_ticker As String
    
    
    
    volume = 0
    high = -0.0001
    greatest_decrease = 0
    total = 0
    
    For Each ws In Worksheets
    
    ws.Cells(1, OUTPUT_TICKER_COL).Value = "ticker"
    ws.Cells(1, OUTPUT_CHANGE_COL).Value = "yearly change"
    ws.Cells(1, OUTPUT_PERCENT_COL).Value = "percent change"
    ws.Cells(1, OUTPUT_TOTAL_COL).Value = "total stock volume"
    
    output_row = FIRST_DATA_ROW
    
    opening_price = ws.Cells(2, 3)
    
    
    last_row = ws.Cells(Rows.Count, INPUT_TICKER_COL).End(xlUp).Row
    
    For input_row = FIRST_DATA_ROW To last_row
        current_ticker = ws.Cells(input_row, INPUT_TICKER_COL).Value
        next_ticker = ws.Cells(input_row + 1, INPUT_TICKER_COL).Value
        
    
    volume = volume + ws.Cells(input_row, 7).Value
        
        If current_ticker <> next_ticker Then
            ' inputs
            closing_price = ws.Cells(input_row, 6).Value
            

            ' calculations
            yearly_change = closing_price - opening_price
            percent_change = yearly_change / opening_price
            
            
            ' outputs
            ws.Cells(output_row, OUTPUT_TICKER_COL).Value = current_ticker
            With ws.Cells(output_row, OUTPUT_CHANGE_COL)
                .Value = yearly_change
                If yearly_change >= 0 Then
                    .Interior.ColorIndex = 4
                    
                Else
                    .Interior.ColorIndex = 3
                    
                    End If
                End With
                
            With ws.Cells(output_row, OUTPUT_PERCENT_COL)
                .Value = percent_change
                .NumberFormat = "0.00%"
                
                If percent_change >= 0 Then
                    .Interior.ColorIndex = 4
                    
                Else
                    .Interior.ColorIndex = 3
                    
                    End If
                    
            End With
            ws.Cells(output_row, OUTPUT_TOTAL_COL).Value = volume
            
            ' prepare for next stock
            output_row = output_row + 1
            If percent_change < greatest_decrease Then
                greatest_decrease = percent_change
                low_ticker = current_ticker
                
            End If
            
            opening_price = ws.Cells(input_row + 1, 3).Value
        
            If percent_change > high Then
                high = percent_change
                high_ticker = current_ticker
                
            End If

            
            If volume > total Then
                total = volume
                total_ticker = current_ticker
            
            End If
            
             volume = 0
             
             
        End If
        
        
    Next input_row
    
    ws.Cells(2, "O").Value = "Greatest % Increase"
    ws.Cells(3, "O").Value = "Greatest % Decrease"
    ws.Cells(4, "O").Value = "Greatest Total Volume"
    ws.Cells(1, "P").Value = "Ticker"
    ws.Cells(1, "Q").Value = "Value"
    ws.Cells(2, "P").Value = high_ticker
    ws.Cells(3, "P").Value = low_ticker
    ws.Cells(4, "P").Value = total_ticker
    With ws.Cells(2, "Q")
        .Value = high
        .NumberFormat = "0.00%"
    End With
    With ws.Cells(3, "Q")
        .Value = greatest_decrease
        .NumberFormat = "0.00%"
    End With
    ws.Cells(4, "Q").Value = total
    
    Next ws
    
End Sub
