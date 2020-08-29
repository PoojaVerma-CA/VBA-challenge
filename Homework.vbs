'-----------------------------
'   Loop through all worksheets
'   Build Summary table 1
'   Build Summary table 2
'-----------------------------


Sub ticker()

' Set initial variables to build Summary Table 1....
Dim ticker As String
Dim Total_stock_vol As Double
Total_stock_vol = 0
'End of Year Closing Price
Dim EOY_closing_price As Double
EOY_closing_price = 0
'Begining of Year Open Price
Dim BOY_open_price As Double
BOY_open_price = 0
'Yearly change
Dim Yearly_change As Double
Yearly_change = 0
'Percent change
Dim Percent_change As Double
Percent_change = 0
  
  
'Initialize variables to build Summary Table 2 ...
'Max Percent Change
Dim Max As Double
'Ticker with Max Percent change
Dim MaxTicker As String
'Min Percent change
Dim Min As Double
'Ticker with Min Percent change
Dim MinTicker As String
'Max stock volume
Dim MaxVol As Double
'Ticker with max volume
Dim MaxVolTicker As String
  
  
    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    For Each ws In Worksheets
     
        'Write the Summary Table 1 headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        'format cells to autofit
        ws.Range("I1:J1").Columns.AutoFit
                        
        ' Determine the Last Row of the first column
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
                
        ' To maintain row count for Summary Table 1
        Dim Summ_Row As Integer
        Summ_Row = 2
        
  
        ' ------------------------------------------------
        ' LOOP THROUGH ALL THE TICKERS IN THE SPREADSHEET
        ' ------------------------------------------------
        For i = 2 To LastRow
            
            ' If this is first row for the ticker, find opening price at the beginning of the year
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                BOY_open_price = ws.Cells(i, 3).Value
            End If

            ' If this is the last row for the ticker, then...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                ' Set the ticker value
                ticker = ws.Cells(i, 1).Value
                
                 ' Find the closing price at the end of that year
                EOY_closing_price = ws.Cells(i, 6).Value

                ' Add to the Total Stock Volume
                Total_stock_vol = Total_stock_vol + ws.Cells(i, 7).Value
     
                'Calculate the Yearly change in price
                Yearly_change = EOY_closing_price - BOY_open_price
      
                'Calculate Perecnt Change
                If BOY_open_price <> 0 Then
                    Percent_change = (Yearly_change / BOY_open_price)
                Else
                    Percent_change = 0
                End If
      

                ' Print the Summary Table 1
                ws.Cells(Summ_Row, 9).Value = ticker
                ws.Cells(Summ_Row, 10).Value = Yearly_change
                ws.Cells(Summ_Row, 11).Value = Format(Percent_change, "0.00%")
                ws.Cells(Summ_Row, 12).Value = Total_stock_vol

                ' Apply conditional formatting on Yearly Change column
                If Yearly_change < 0 Then
                    ws.Cells(Summ_Row, 10).Interior.ColorIndex = 3
                ElseIf Yearly_change > 0 Then
                    ws.Cells(Summ_Row, 10).Interior.ColorIndex = 4
                End If
      
                ' Increment row for Summary Table 1
                Summ_Row = Summ_Row + 1
      
                ' Reset the Total_stock_vol
                Total_stock_vol = 0

            ' If the cell immediately following a row is the same ticker...
            Else
                ' Add to the Brand Total
                Total_stock_vol = Total_stock_vol + ws.Cells(i, 7).Value

            End If
    
        Next i
        'format Summary Table 1 cells to autofit
        ws.Range("I:L").Columns.AutoFit
        
    
' --------------------------------------------
' Summary table 1 completed
' --------------------------------------------
            

        
        'set the starting max and min values
        Max = ws.Cells(2, 11).Value
        Min = ws.Cells(2, 11).Value
        MaxVol = ws.Cells(2, 12).Value
        
        ' ------------------------------------------------------------
        ' Now loop through Summary Table 1 to build Summary Table 2
        ' ------------------------------------------------------------
        
        'Find last row of summary table
        SummTableLastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
                
        For i = 2 To SummTableLastRow
            
            'Find max and Min Percent change
            If ws.Cells(i, 11).Value > Max Then
                Max = ws.Cells(i, 11).Value
                MaxTicker = ws.Cells(i, 9)
            ElseIf ws.Cells(i, 11).Value < Min Then
                Min = ws.Cells(i, 11).Value
                MinTicker = ws.Cells(i, 9)
            End If
            
            'Find Max Volume
            If ws.Cells(i, 12).Value > MaxVol Then
                MaxVol = ws.Cells(i, 12).Value
                MaxVolTicker = ws.Cells(i, 9)
            End If
            
        Next i
        
        'Print the Challenge Summary table
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(2, 16).Value = MaxTicker
        ws.Cells(2, 17).Value = Max
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(3, 16).Value = MinTicker
        ws.Cells(3, 17).Value = Min
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(4, 16).Value = MaxVolTicker
        ws.Cells(4, 17).Value = MaxVol
        ws.Range("Q4").NumberFormat = "0.0000E+00"
        
       'format cells in to autofit
        ws.Range("O1:Q4").Columns.AutoFit
    
    Next ws

End Sub

