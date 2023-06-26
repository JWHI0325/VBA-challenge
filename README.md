'VBA-challenge
'Module 2 

Sub ticker_stock()
    'create worksheet
    Dim ws As Worksheet
    
    'label our variables
        Dim ticker As String
        Dim Volume As Double
        Dim Open_Price As Double
        Dim Close_Price As Double
        Dim Summary_table As Integer
        Dim Year_open_row As Long
        Dim last_row As Long
        Dim i As Long
        Dim Worksheetname As String
        Dim Row As Double
        Dim Yearly_Change As Double
        Dim Ticker_Row As Long
        Dim Percent_change As Double
        
     For Each ws In Worksheets
      
        'setting values
        TickerRow = 2
        Summary_table = 2
        Volume = 0
        Year_open_row = 2
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        Open_Price = 0
        Close_Price = 0
        Row = 2
        
        'create headers

        ws.Cells(1, 9).Value = "ticker"
        ws.Cells(1, 10).Value = "Yearly_Change"
        ws.Cells(1, 11).Value = "Percent_Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
       
        
        'start loop
        
        For i = 2 To last_row
        
        'define open price and volume
        Open_Price = ws.Cells(Year_open_row, 3).Value
        Volume = Volume + ws.Cells(i, 7).Value
        ticker = ws.Cells(i, 1).Value
        
        
        'conditional for loop
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
                        
            'setting ticker name
                       
            ws.Cells(TickerRow, 9).Value = ticker
            
        
        
           'set close prices
            Close_Price = ws.Cells(i, 6).Value
           
            
            'add yearly change
            
         Yearly_Change = Close_Price - Open_Price
        
        
            ws.Cells(TickerRow, 10).Value = Yearly_Change
            
            'add conditional
            If Open_Price = 0 Then
                ws.Cells(TickerRow, 11).Value = Null
                          
            Else
                Percent_change = Yearly_Change / Open_Price
                ws.Cells(TickerRow, 11).Value = Percent_change
                ws.Cells(TickerRow, 11).NumberFormat = "0.00%"
                
              
           End If
            'add total volume
            
             ws.Cells(TickerRow, 12).Value = Volume
            Volume = 0
            TickerRow = TickerRow + 1
            Year_open_row = i + 1
            
            End If
            
    Next i
        'assign the value to cells
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Value"
        ws.Cells(1, 16).Value = "ticker"
        ws.Cells(1, 17).Value = "Value"
        
        summary_table_last_row = ws.Cells(Rows.Count, 11).End(xlUp).Row
        
     
    greatest_percent_increase = WorksheetFunction.Max(Range("K2:K" & summary_table_last_row))
    ws.Cells(2, 17).Value = greatest_percent_increase
    ws.Cells(2, 17).NumberFormat = "0.00%"
    greatest_percent_increase_ticker = ws.Cells(Application.WorksheetFunction.Match(greatest_percent_increase, Range("K2:K" & summary_table_last_row), 0) + 1, 9).Value
    ws.Cells(2, 16).Value = greatest_percent_increase_ticker
    
    
   greatest_percent_decrease = WorksheetFunction.Min(Range("K2:K" & summary_table_last_row))
    ws.Cells(3, 17).Value = greatest_percent_decrease
    ws.Cells(3, 17).NumberFormat = "0.00%"
   greatest_percent_decrease_ticker = ws.Cells(Application.WorksheetFunction.Match(greatest_percent_decrease, Range("K2:K" & summary_table_last_row), 0) + 1, 9).Value
    ws.Cells(3, 16).Value = greatest_percent_decrease_ticker
   

    greatest_total_value = WorksheetFunction.Max(Range("L2:L" & summary_table_last_row))
    ws.Cells(4, 17).Value = greatest_total_value
    greatest_total_value_ticker = ws.Cells(Application.WorksheetFunction.Match(greatest_total_value, Range("L2:L" & summary_table_last_row), 0) + 1, 9).Value
    ws.Cells(4, 16).Value = greatest_total_value_ticker
    
    
    'conditional formatting loop
    If Yearly_Change > 0 Then
        ws.Cells(Summary_table, 10).Interior.ColorIndex = 3
        
    ElseIf Yearly_Change < 0 Then
        ws.Cells(Summary_table, 10).Interior.ColorIndex = 4
        
    End If
    
    Summary_table = Summary_table + 1
    
    Next
    
            
    End Sub
    
    

