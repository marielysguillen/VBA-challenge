Sub stock_years()

'-------------------Variable declarations:-------------------------

Dim current_ticker_a As String
Dim yearly_change_j As Double
Dim percent_change_k As Double
Dim percent_change_k_string As String
Dim close_price_f As Double
Dim open_price_c As Double
Dim ticker_name_a As String

Dim greatest_increase_q As Double
Dim greatest_decrease_q As Double
Dim greatest_total_vol_q As Double

Dim increase_total_ticker_p As String
Dim decrease_total_ticker_p As String
Dim greatest_total_vol_ticker_p As String

Dim total_stock_volume_l As Double
total_stock_volume_l = 0

' Keep track of the tickers in the summary table
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

'to enable it to run on every worksheet (that is, every year) at once.
Dim ws As Worksheet


' --------------------------------------------
' LOOP THROUGH ALL SHEETS to run on every worksheet at once
' --------------------------------------------
   For Each ws In Worksheets
   
   'Clear variables - New sheet
     yearly_change_j = 0
     percent_change_k = 0
     total_stock_volume_l = 0
     close_price_f = 0
     open_price_c = 0
     
     greatest_increase_q = 0
     greatest_decrease_q = 0
     greatest_total_vol_q = 0
     
     'Find tickers
     Summary_Table_Row = 2
     
     ' Determine the Last Row
     LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
     'Determine opening price
     open_price_c = ws.Cells(2, 3).Value
    
    
'               ------------------- Create Header -------------------------
    'Header for the following output:
     '-The ticker symbol
     '-Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
     '-The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
     '-The total stock volume of the stock
     
     ws.Range("I1").Value = "Ticker"
     ws.Range("J1").Value = "Yearly Change"
     ws.Range("K1").Value = "Percent Change"
     ws.Range("L1").Value = "Total Stock Volume"
     
     'Header to return the stock with the
     '-"Greatest % increase"
     '-"Greatest % decrease"
     '-"Greatest total volume"
     ws.Range("P1").Value = "Ticker"
     ws.Range("Q1").Value = "Value"
     ws.Range("O2").Value = "Greatest % Increase"
     ws.Range("o3").Value = "Greatest % Decrease"
     ws.Range("O4").Value = "Greatest Total Volume"
     
     
' --------------------------------------------
' LOOP THROUGH the current sheet
' --------------------------------------------
    'Loop through all stocks sheets
        For i = 2 To LastRow
        
       'Count
       'row_one = row_one + 1
            
            'Let's check if we are still within the same ticker.. if we are not within the same ticker...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
              'Set current ticket name
                current_ticker = ws.Cells(i, 1).Value
                'MsgBox (current_ticker)
                
'------------------- Set cells in the spreadsheet for columns I,J,K,L -------------------------

                'Add Total Stock Volume for the current row
                total_stock_volume_l = total_stock_volume_l + ws.Cells(i, 7).Value
              
                'Add closing price for the current row
                close_price_f = ws.Cells(i, 6).Value
                
                yearly_change_j = (close_price_f - open_price_c)
                
                 If open_price_c <> 0 Then 'Avoid error
                  percent_change_k = (yearly_change_j / open_price_c) * 100
                 End If
                 
                 ' Print the Ticker in the Summary Table
                 ws.Range("I" & Summary_Table_Row).Value = current_ticker
                 
                 ' Print the Yearly change to the Summary Table
                 ws.Range("J" & Summary_Table_Row).Value = yearly_change_j
                 
                 ' Print the Percent change to the Summary Table
                  percent_change_k_string = (Str(percent_change_k) & "%")
                  ws.Range("K" & Summary_Table_Row).Value = percent_change_k_string
                  percent_change_k_string = ""
                 
                 
                 ' Print the Total stick volume to the Summary Table
                 ws.Range("L" & Summary_Table_Row).Value = total_stock_volume_l
                 
                 'Color column J
                 If (yearly_change_j > 0) Then
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                 ElseIf (yearly_change_j <= 0) Then
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                 End If


                 ' Add one to the summary table row
                 Summary_Table_Row = Summary_Table_Row + 1
        
                 'Clear variables
                 yearly_change_j = 0
                 percent_change_k = 0
                 total_stock_volume_l = 0
                 close_price_f = 0
                 
                 'Reset opening price
                 open_price_c = ws.Cells(i + 1, 3).Value
           
            
            Else
                 
             'total_stock_volume Results column 7
              total_stock_volume_l = total_stock_volume_l + ws.Cells(i, 7).Value
            
                
            End If 'End of we are still within the same ticker
    
        Next i 'End of Loop2
    
    'Set greatest initial values row 2
      greatest_increase_q = Cells(2, 11)
      greatest_decrease_q = Cells(2, 11)
      greatest_total_vol_q = Cells(2, 12)
     
     
    ' --------------------------------------------
    ' LOOP THROUGH the current sheet
    ' --------------------------------------------
     For j = 2 To LastRow
    'return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume
    
       If Cells(j, 11) > greatest_increase_q Then
        greatest_increase_q = Cells(j, 11)
        Cells(2, 17) = greatest_increase_q
        Cells(2, 17).NumberFormat = "0.00%"
        Cells(2, 16) = Cells(j, 9)
       End If


        If Cells(j, 11) < greatest_decrease_q Then
        greatest_decrease_q = Cells(j, 11)
        Cells(3, 17) = greatest_decrease_q
        Cells(3, 17).NumberFormat = "0.00%"
        Cells(3, 16) = Cells(j, 9)
  
        End If
      
        If Cells(j, 12) > greatest_total_vol_q Then
        greatest_total_vol_q = Cells(j, 12)
        Cells(4, 17) = greatest_total_vol_q
        Cells(4, 16) = Cells(j, 9)
   
        End If
   
    Next j 'End of Loop3

    Next ws 'End of Loop1

'    MsgBox ("Fixes Complete")

End Sub



