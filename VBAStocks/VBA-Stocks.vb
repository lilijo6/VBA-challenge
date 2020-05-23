'Create a script that will loop through all the stocks for one year and output the following information.
' The ticker symbol.
'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
'The total stock volume of the stock.
'You should also have conditional formatting that will highlight positive change in green and negative change in red.

Sub StockData()
    
    'Make the appropriate adjustments to your VBA script that will allow it to run on every worksheet, i.e., every year, just by running the VBA script once.
    'All Cell and range values were modified to contain a "ws." in front of it
    For Each ws In Worksheets
   
        ' INSERT THE TICKER SYMBOL, YEARLY CHANGE AND PERCENT CHANGE

        ' Created a Variables
        Dim Ticker_Symbol As String
        Dim Yearly_Change As Double
        Dim Percent_change As Double
        Dim Tot_Stock As Double
        Dim open_price As Double
        Dim close_price As Double
        Tot_Stock = 0
        
        'Keep track of the location of each Ticker Symbol in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
        'Create Summary Table
        ws.Cells(1, 11) = "Ticker"
        ws.Cells(1, 12) = "Yearly Change"
        ws.Cells(1, 13) = "Percent Change"
        ws.Cells(1, 14) = "Total Stock Volume"
        ws.Columns(12).NumberFormat = "0.00"
        ws.Columns(13).NumberFormat = "0.00%"
        
        'Define open price
        open_price = ws.Cells(2, 3).Value
        
        ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
            'Loop through all Tickers
            For i = 2 To LastRow
            
                'Check if we are still within the same Ticker"
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                                    
                    'Grab/Set the Ticker Symbol
                    Ticker_Symbol = ws.Cells(i, 1).Value
                                        
                    'Add to the Total Stock Volume
                    Tot_Stock = Tot_Stock + ws.Cells(i, 7).Value
                    
                    'Print the Ticker Symbol in the Summary Table
                    ws.Range("K" & Summary_Table_Row).Value = Ticker_Symbol
             
                    'Print the Total Stock Volume to the Summary Table
                    ws.Range("N" & Summary_Table_Row).Value = Tot_Stock
                    
                    'Add one to the summary table row
                    Summary_Table_Row = Summary_Table_Row + 1
                    
                    'Reset the Total Stock Volume
                    Tot_Stock = 0
                    
                    'Calculating the yearly change
                    close_price = ws.Cells(i, 6).Value
                    Yearly_Change = close_price - open_price
                
                                        
                    'Print the Yearly Change
                    ws.Range("L" & Summary_Table_Row - 1).Value = Yearly_Change
                                       
                                                      
                    'Calculating the Percent change
                    If open_price <> 0 And close_price <> 0 Then
                    Percent_change = (close_price / open_price) - 1
                    open_price = ws.Cells(i + 1, 3)
                                                          
                    End If
                                                            
                    'Print the Percent Change
                    ws.Range("M" & Summary_Table_Row - 1).Value = Percent_change
                                          
                
                ' If the cell immediately following a row is the same
                Else
                
                'Add to the Total Stock Volume
                Tot_Stock = Tot_Stock + ws.Cells(i, 7).Value
                
                End If
                
                
                
            Next i
            
        'CONDITIONAL FORMATTING
     
      
        ' Determine the Last on Summary Table
        RowCount = ws.Cells(Rows.Count, 12).End(xlUp).Row
        
            'Create a new loop for Yearly_Change
                    
            For j = 2 To RowCount
        
                Yearly_Change = ws.Cells(j, 12).Value
        
                If Yearly_Change >= 0 Then
                  
                ws.Cells(j, 12).Interior.ColorIndex = 4
    
                Else
    
                ws.Cells(j, 12).Interior.ColorIndex = 3
    
                End If
    
            Next j
            
        
 
    Next ws
         
End Sub




