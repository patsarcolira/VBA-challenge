Option Explicit
Sub Aphabetical_testing()
    Dim LastRow As Long
    Dim LastColumn As Long
    Dim A_Row As Long
    Dim Ticker_Counter As Long
    Dim Change As Boolean
    Dim Open_price As Double
    Dim Close_price As Double
    Dim Yearly_change As Double
    Dim Percent_change As Double
    Dim Total_stock_volume As Double
    Dim Max_Increase As Double
    Dim Min_Decrease As Double
    Dim Max_Volume As Double
    

    'Title of reults rows
    Range("I1").Value = "Ticker symbol"
    Range("J1").Value = "Yearly change"
    Range("K1").Value = "Percent change"
    Range("L1").Value = "Total stock volume of the stock"
    
    Range("O1").Value = ""
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    
    
    'Size of data
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    'Debug.Print (LastRow)
    LastColumn = Cells(1, Columns.Count).End(xlUp).Column
    'Debug.Print (LastColumn)
    
    'Inizialize variables
    A_Row = 0
    Ticker_Counter = 0
    Change = True
    Open_price = 0
    Close_price = 0
    Yearly_change = 0
    Percent_change = 0
    Total_stock_volume = 0
    Max_Increase = 0
    Max_Volume = 0
    'Print Ticker
    
    For A_Row = 2 To LastRow
        If Change = True Then
            Open_price = Cells(A_Row, 3).Value
        End If
        Total_stock_volume = Total_stock_volume + Cells(A_Row, 7).Value
        
        If Cells(A_Row, 1).Value <> Cells(A_Row + 1, 1) Then
            Range("I" & 2 + Ticker_Counter).Value = Cells(A_Row, 1).Value
            
            Close_price = Cells(A_Row, 6)
            Yearly_change = Close_price - Open_price
            Range("J" & 2 + Ticker_Counter).Value = Round(Yearly_change, 2)
            
            Percent_change = Yearly_change * 100 / Open_price
            Range("K" & 2 + Ticker_Counter).Value = Round(Percent_change, 2)
            Range("K" & 2 + Ticker_Counter).NumberFormat = "0.00%"
            
            Range("L" & 2 + Ticker_Counter).Value = Round(Total_stock_volume, 2)
            
            'Debug.Print (Total_stock_volume)
            Total_stock_volume = 0
            'Debug.Print (Total_stock_volume)
            Change = True
            Ticker_Counter = Ticker_Counter + 1
           
        Else
            Change = False
        End If
    Next A_Row
    
    LastRow = Range("I1", Range("I1").End(xlDown)).Rows.Count
    
    For A_Row = 2 To LastRow
    If Cells(A_Row, 11).Value > Max_Increase Then
        Max_Increase = Cells(A_Row, 11).Value
       
        Range("Q" & 2).Value = Max_Increase
        Range("Q" & 2).NumberFormat = "0.00%"
        Range("P" & 2).Value = Cells(A_Row, 9).Value
    End If
    If Cells(A_Row, 11).Value < Min_Decrease Then
        Min_Decrease = Cells(A_Row, 11).Value
        Range("Q" & 3).Value = Min_Decrease
        Range("Q" & 3).NumberFormat = "0.00%"
        Range("P" & 3).Value = Cells(A_Row, 9).Value
    End If
    If Cells(A_Row, 12).Value > Max_Volume Then
        Max_Volume = Cells(A_Row, 12).Value
        Range("Q" & 4).Value = Max_Volume
        Range("P" & 4).Value = Cells(A_Row, 9).Value
    End If

    Next A_Row
    
 
 
    
End Sub





