Attribute VB_Name = "Module1"
Option Explicit
Sub Practice()
    Dim LastRow As Long
    Dim LastColumn As Long
    Dim A_Row As Long
    Dim Ticker_Counter As Long


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
    Debug.Print (LastRow)
    LastColumn = Cells(1, Columns.Count).End(xlUp).Column
    Debug.Print (LastColumn)
    A_Row = 0
    Ticker_Counter = 0
    'Print Ticker
    For A_Row = 2 To LastRow
        If Cells(A_Row, 1).Value <> Cells(A_Row + 1, 1) Then
            Range("I" & 2 + Ticker_Counter).Value = Cells(A_Row, 1).Value
            Ticker_Counter = Ticker_Counter + 1
            If Cells(A_Row, 1).Value <> Cells(A_Row + 1, 1) Then
            Range("I" & 2 + Ticker_Counter).Value = Cells(A_Row, 1).Value
            Ticker_Counter = Ticker_Counter + 1
            
        
        End If
    Next A_Row
    
 
    Debug.Print (Ticker_Counter)
    Debug.Print (A_Row)
    Debug.Print ("Hola")
        
   
        
        
    
    
    

    
    
    
    
End Sub





