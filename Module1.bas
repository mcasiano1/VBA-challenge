Attribute VB_Name = "Module1"
Sub StockMarketAnalysis()


Dim ws As Worksheet
Dim starting_ws As Worksheet
Set starting_ws = ActiveSheet

For Each ws In ThisWorkbook.Worksheets
ws.Activate

    'Set an initial variable for holding the ticker symbol
    Dim ticker_symbol As String
    'Set variables to calculate Yearly Change
    Dim opening_price As Double
    Dim closing_price As Double
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    
    'Set Column Labels on Sheet
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    'Set an initial variable for holding the Total Stock Volume per ticker
    Dim total_stock_vol As Double
    total_stock_vol = 0
    
    'Keep track of the location for each ticker symbol in the column summary
    Dim ticker_column_summary As Double
    ticker_column_summary = 2
    
    'Loop through all ticker values
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastrow
    
        'Check if we are still in the same ticker value, if it's not...
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
         'Set the ticker value
             ticker_symbol = Cells(i, 1).Value
             
             'Set the closing price and opening price
              closing_price = Cells(i, 6).Value
              
                'Calculate the Yearly Change and Percent Change
                 Yearly_Change = closing_price - opening_price
                 
            'Add to the Total Stock Volume
             total_stock_vol = total_stock_vol + Cells(i, 7).Value
             
             'get opening price for next ticker
            If (opening_price <> 0) Then
            
            'Calculate Percent Change
                Percent_Change = ((closing_price - opening_price) / opening_price)
                
             Else
                
                Percent_Change = 0
              
            End If
            
            Range("K" & ticker_column_summary).NumberFormat = "0.00%"
           
        'Print the ticker symbol in the Summary Column
         Range("I" & ticker_column_summary).Value = ticker_symbol
         'Print the Yearly Change to the Summary Column
         Range("J" & ticker_column_summary).Value = Yearly_Change
         'Print the Percent Change to the Summary Column
         Range("K" & ticker_column_summary).Value = Percent_Change
         
         'Format cells with color
         If (Percent_Change >= 0) Then
         Range("J" & ticker_column_summary).Interior.ColorIndex = 4
         
         Else
         Range("J" & ticker_column_summary).Interior.ColorIndex = 3
         
         End If
         
        'Print the Total Stock Volume to the Summary Column
         Range("L" & ticker_column_summary).Value = total_stock_vol
        'Add one to the  ticker column summary
         ticker_column_summary = ticker_column_summary + 1
        'Reset the Total Stock Volume
         total_stock_vol = 0
         opening_price = Cells(i + 1, 3).Value
    'If the cell immediatley following a row is the same ticker symbol...
    Else
        'Set opening price
       If i = 2 Then
           opening_price = Cells(2, 3).Value
       End If
        'Add to the Total Stock Volume
        total_stock_vol = total_stock_vol + Cells(i, 7).Value
     
        
    End If
       
  Next i
  
  ws.Cells(1, 1) = 1
  
  Next
  
  starting_ws.Activate
  
    
End Sub

