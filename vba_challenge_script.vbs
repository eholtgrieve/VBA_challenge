Attribute VB_Name = "Module1"
Sub ticker_counter(worksheet_number)
    
    'setting variables
    
    Dim ticker_name As String
    Dim yearly_change As Double
    Dim stock_volume As Double
    Dim previous_amount As Double
    Dim year_open As Double
    Dim year_close As Double
    Dim percent_change As Double
    
    Dim greatest_increase As Double
    Dim greatest_decrease As Double
    Dim greatest_volume As Double
    
    yearly_change = 0
    stock_volume = 0
    previous_amount = 2
    greatest_increase = 0
    greatest_decrease = 0
    greatest_volume = 0
    
    Dim ticker_table_row As Integer
    ticker_table_row = 2
    
    'adding columns
    ActiveWorkbook.Worksheets(worksheet_number).Columns("I:P").EntireColumn.Insert
    ActiveWorkbook.Worksheets(worksheet_number).Cells(1, 9).Value = "Ticker"
    ActiveWorkbook.Worksheets(worksheet_number).Cells(1, 10).Value = "Yearly Change"
    ActiveWorkbook.Worksheets(worksheet_number).Cells(1, 11).Value = "Percent Change"
    ActiveWorkbook.Worksheets(worksheet_number).Cells(1, 12).Value = "Total Stock Volume"
    
    ActiveWorkbook.Worksheets(worksheet_number).Cells(1, 15).Value = "Ticker"
    ActiveWorkbook.Worksheets(worksheet_number).Cells(1, 16).Value = "Value"
    ActiveWorkbook.Worksheets(worksheet_number).Cells(2, 14).Value = "Greatest % Increase"
    ActiveWorkbook.Worksheets(worksheet_number).Cells(3, 14).Value = "Greatest % Decrease"
    ActiveWorkbook.Worksheets(worksheet_number).Cells(4, 14).Value = "Greatest Total Volume"
    
    
    
    'loop loop loop
 
    For i = 2 To 753001
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            ticker_name = Cells(i, 1).Value
            year_open = Range("C" & previous_amount)
            year_close = Range("F" & i)
            
            yearly_change = year_close - year_open
            
            If year_open = 0 Then
                percent_change = 0
            Else
                percent_change = yearly_change / year_open
            End If
            
            stock_volume = stock_volume + Cells(i, 7)
            
            Range("I" & ticker_table_row).Value = ticker_name
            Range("J" & ticker_table_row).Value = yearly_change
            Range("L" & ticker_table_row).Value = stock_volume
            Range("K" & ticker_table_row).Value = percent_change
            Range("K" & ticker_table_row).NumberFormat = "0.00%"
           
            
            If Range("K" & ticker_table_row).Value > greatest_increase Then
                Range("O2").Value = ticker_name
                Range("P2").Value = percent_change
                Range("P2").NumberFormat = "0.00%"
            End If
            
            If Range("K" & ticker_table_row).Value < greatest_decrease Then
                Range("O3").Value = ticker_name
                Range("P3").Value = percent_change
                Range("P3").NumberFormat = "0.00%"
            
            End If
        
            If Range("L" & ticker_table_row).Value > greatest_volume Then
                Range("O4").Value = ticker_name
                Range("P4").Value = stock_volume
           
            
            End If
            
            
            If Range("J" & ticker_table_row).Value > 0 And Range("K" & ticker_table_row).Value > 0 Then
                Range("J" & ticker_table_row).Interior.ColorIndex = 4
                Range("K" & ticker_table_row).Interior.ColorIndex = 4
            Else
                Range("J" & ticker_table_row).Interior.ColorIndex = 3
                Range("K" & ticker_table_row).Interior.ColorIndex = 3
            End If
            
            ticker_table_row = ticker_table_row + 1
            previous_amount = i + 1
            
            stock_volume = 0
            yearly_change = 0
            percent_change = 0
            
            
        
        Else
            yearly_change = year_close - year_open
            stock_volume = stock_volume + Cells(i, 7)
        End If
        
        
    
    Next i
        
    
            
    
End Sub
Sub process_all_worksheets()
 Dim number_of_worksheets As Integer
 number_of_worksheets = ActiveWorkbook.Worksheets.Count
 Dim worksheet_number As Integer
 
    For worksheet_number = 1 To number_of_worksheets
        ActiveWorkbook.Worksheets(worksheet_number).Select
        Call ticker_counter(worksheet_number)
        ActiveWorkbook.Worksheets(worksheet_number).Columns("I:P").AutoFit
    Next worksheet_number
    
End Sub
