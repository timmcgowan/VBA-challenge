Sub StockAnalysis()
    'define some variables
    Dim tick As String
    Dim price_chg As Double
    Dim percent_chg As Double
    Dim stock_vol As Double
    Dim mark As String
    'define some formating
    mark = "A2" 'set reference point
    output = "I2" 'set output reference
    'Create a script that will loop through all the stocks for one year and output the following information.
    
    'First line of data
    Debug.Print "Starting mark: "; mark
    Range(mark).Select
    
    'Process rows
    'Do Until IsEmpty(ActiveCell) Or ActiveCell.Address(0, 0) = "A300" 'debug only
    Do Until IsEmpty(ActiveCell)
        'check ticker
        If ActiveCell.Value <> ActiveCell.Offset(1, 0).Value Then
            Debug.Print "Current Ticker :"; ActiveCell.Value
            Debug.Print "Previous Ticker :"; Range(mark).Value
            Debug.Print "mark price"; Range(mark).Value
            'Do some math
            'price change in year
            price_chg = Range(mark).Offset(0, 5).Value - ActiveCell.Offset(0, 5).Value
            Debug.Print "Price Change "; price_chg
            
            'percent change
            percent_chg = Round((price_chg / Range(mark).Offset(0, 5).Value * 100), 2)
            Debug.Print "Percent Change "; percent_chg
                        
            'total volume
            stock_vol = stock_vol + ActiveCell.Offset(0, 6).Value
            Debug.Print "Total is "; Total
                        
            'write output
            'The ticker symbol.
            'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
            'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
            'The total stock volume of the stock.
            
            Debug.Print "Output mark: "; output
            Range(output).Offset(0, 0).Value = Range(mark).Value
            Range(output).Offset(0, 1).Value = price_chg
            Range(output).Offset(0, 2).Value = percent_chg
            Range(output).Offset(0, 3).Value = stock_volume
            
            'update mark
            mark = ActiveCell.Offset(1, 0).Address(0, 0)
            Debug.Print "New Mark"; mark
            
        Else
            Total = Total + ActiveCell.Offset(0, 5).Value
        End If
        'Next
        ActiveCell.Offset(1, 0).Select
    Loop
    
        'You should also have conditional formatting that will highlight positive change in green and negative change in red.
    ' Ran out of time, will add later.
    
    'The result should look as follows.
End Sub

