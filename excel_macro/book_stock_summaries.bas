Attribute VB_Name = "Module21"
Sub multiple_year_stock_summary():
             
     For Each leaf In Worksheets
             
        ' declare variables
        Dim ticker_cell, date_cell As String
        Dim x, counter, counter_two, counter_three, counter_four As Integer
        Dim last_row, last_column As Long
        Dim volume, total_volume, end_volume As Double
        Dim open_price, close_price, percent_change, yearly_change, end_high_percent, end_low_percent As Double
        
        ' declare arrays
        Dim ticker() As String
        Dim year_open() As Single
        Dim year_close() As Single
        Dim volume_sum() As Double
        
        
        '|------------------------|
        '|     MODERATE MODE      |
        '|------------------------|
        
        ' define end of rows and columns
        last_row = leaf.Rows.End(xlDown).Row
        last_column = leaf.Columns.End(xlToRight).Column
        
        ' initiate counters
        counter = 0
        counter_two = 0
        counter_three = 0
        counter_four = 0
        
        
        ' new table headers
        If leaf.Cells(1, last_column + 2).Value <> "Ticker_Symbol" Then
            
            leaf.Cells(1, last_column + 2).Value = "Ticker_Symbol"
            leaf.Cells(1, last_column + 3).Value = "Yearly_Change"
            leaf.Cells(1, last_column + 4).Value = "Percent_Change"
            leaf.Cells(1, last_column + 5).Value = "Total_Stock_Volume"
        
        End If
        
        
        ' loop through table to load arrays
        For i = 2 To last_row
                 
            ' initiate load-loop variables
            ticker_cell = leaf.Cells(i, 1).Value
            open_price = leaf.Cells(i, 3).Value
            close_price = leaf.Cells(i, 6).Value
                    
                
            ' load ticker and year_close arrays
            If leaf.Cells(i + 1, 1).Value <> ticker_cell Then
            
                ' re-dim arrays to maintain integrity
                ReDim Preserve ticker(counter)
                ReDim Preserve year_close(counter)
                    
                ' loading arrays
                ticker(counter) = ticker_cell
                year_close(counter) = close_price
                
                counter = counter + 1
            
            End If
            
            
            ' load year_open array
            If leaf.Cells(i - 1, 1).Value <> ticker_cell Then
            
                ' re-dim arrays to maintain integrity
                ReDim Preserve year_open(counter_two)
                
                'loading arrays
                year_open(counter_two) = open_price
                
                counter_two = counter_two + 1
            
            End If
              
        Next i
        
        
        ' for loop with embedding while loop to load volume_sum array
        x = 2
        For j = 2 To counter + 1
            
            total_volume = 0
            Do While leaf.Cells(x, 1) = ticker(j - 2)
                
                volume = leaf.Cells(x, 7).Value
            
                x = x + 1
                total_volume = total_volume + volume
            
            Loop
            
            ' rebuilding volume_sum array integrity
            ReDim Preserve volume_sum(counter_three)
            
            ' loading volum_sum array
            volume_sum(counter_three) = total_volume
            counter_three = counter_three + 1
        
        Next j
    
        
        ' loop through new table to unload values from arrays
        For k = 2 To counter + 1
            
            ' calculations for yearly_change and percent_change columns
            yearly_change = year_close(counter_four) - year_open(counter_four)
            
            If year_open(counter_four) <> 0 Then
                percent_change = year_close(counter_four) / year_open(counter_four) - 1
            Else
                percent_change = 0
            End If
            
            ' unload summary columns
            leaf.Cells(k, last_column + 2).Value = ticker(counter_four)
            leaf.Cells(k, last_column + 3).Value = yearly_change
            leaf.Cells(k, last_column + 4).Value = percent_change
            leaf.Cells(k, last_column + 5).Value = volume_sum(counter_four)
            
            ' format yearly_change and percent_change columns
            If yearly_change > 0 Then
                
                leaf.Cells(k, last_column + 3).Interior.Color = RGB(0, 250, 0)
            ElseIf yearly_change = 0 Then
                leaf.Cells(k, last_column + 3).Interior.Color = RGB(255, 255, 255)
            Else
                leaf.Cells(k, last_column + 3).Interior.Color = RGB(250, 0, 0)
            
            End If
            
            leaf.Cells(k, last_column + 3).NumberFormat = "0.00"
            leaf.Cells(k, last_column + 4).NumberFormat = "0.00%"
            
            counter_four = counter_four + 1
            
        Next k
        
        
        '|------------------------|
        '|       HARD MODE        |
        '|------------------------|
        
        ' insert column and row labels
        leaf.Cells(1, last_column + 8).Value = "Ticker"
        leaf.Cells(1, last_column + 9).Value = "Value"
        
        leaf.Cells(2, last_column + 7).Value = "Greatest % Increase"
        leaf.Cells(3, last_column + 7).Value = "Greatest % Decrease"
        leaf.Cells(4, last_column + 7).Value = "Greatest Total Volume"
        
        
        ' initiate column variable
        end_high_percent = leaf.Cells(2, last_column + 4).Value
        end_low_percent = leaf.Cells(2, last_column + 4).Value
        end_volume = leaf.Cells(2, last_column + 5).Value
        
        
        ' fill cell values for end summary
        For l = 2 To counter + 1
            
            ' highest value increase in yearly_change column
            If leaf.Cells(l + 1, last_column + 4).Value > end_high_percent Then
                leaf.Cells(2, last_column + 8).Value = leaf.Cells(l + 1, last_column + 2).Value
                end_high_percent = leaf.Cells(l + 1, last_column + 4).Value
                leaf.Cells(2, last_column + 9) = end_high_percent
                leaf.Cells(2, last_column + 9).NumberFormat = "0.00%"
            End If
            
            ' highest value decrease in yearly_change column
            If leaf.Cells(l + 1, last_column + 4).Value < end_low_percent Then
                leaf.Cells(3, last_column + 8).Value = leaf.Cells(l + 1, last_column + 2).Value
                end_low_percent = leaf.Cells(l + 1, last_column + 4).Value
                leaf.Cells(3, last_column + 9) = end_low_percent
                leaf.Cells(3, last_column + 9).NumberFormat = "0.00%"
            End If
            
            ' greatest total in total_stock_volume column
            If leaf.Cells(l + 1, last_column + 5).Value > end_volume Then
                leaf.Cells(4, last_column + 8).Value = leaf.Cells(l + 1, last_column + 2).Value
                end_volume = leaf.Cells(l + 1, last_column + 5).Value
                leaf.Cells(4, last_column + 9) = end_volume
            End If
            
            ' refresh column variables with new values
            end_high_percent = end_high_percent
            end_low_percent = end_low_percent
            end_volume = end_volume
            
        Next l
    
    Next
    MsgBox ("Workbook summary complete, have a nice day :)")
        
End Sub
