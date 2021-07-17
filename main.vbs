Sub wall_street():

    Dim LastRow As Long
    Dim LastColumn As Long

    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Debug.Print (LastRow)
    LastColumn = Cells(1, Columns.Count).End(xlUp).Column
    Debug.Print (LastColumn)
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
    
    
    
    
    

end Sub
