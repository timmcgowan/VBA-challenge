Sub StockAnalysis()
    'define some variables
    Dim price_chg As Double
    Dim percent_chg As Double
    Dim stock_vol As Double
    Dim mark As String
    Dim output As String
           
    'define some formating
    mark = "A2" 'set reference point
    output = "I2" 'set output reference
    Dim rng, frng As Range
    
    Dim git, gip, gdt, gdp, gvt, gvl As Variant
        
    'clear formating
    Set frng = Union(Range(output), Range(output).Offset(0, 2))
    frng.FormatConditions.Delete
            
    'Create a script that will loop through all the stocks for one year and output the following information.
            
    'First line of data
    Debug.Print "Starting mark: "; mark
    Range(mark).Select
    
    'Process rows
    'Do Until IsEmpty(ActiveCell) Or ActiveCell.Address(0, 0) = "A300" 'debug only
    Do Until IsEmpty(ActiveCell)
        'check ticker
        If ActiveCell.Value <> ActiveCell.Offset(1, 0).Value Then
            Debug.Print "Current Ticker :"; Range(mark).Value
            Debug.Print "Next Ticker :"; ActiveCell.Offset(1, 0).Value
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
            Range(output).Value = Range(mark).Value
            Range(output).Offset(0, 1).Value = price_chg
            If price_chg > gip Then
                git = Range(mark).Value
                gip = price_chg
            End If
            If price_chg > gdp Then
                gdt = Range(mark).Value
                gdp = price_chg
            End If
            Range(output).Offset(0, 2).Value = percent_chg

            Range(output).Offset(0, 3).Value = stock_vol
            If stock_vol > gvl Then
                gvt = Range(mark).Value
                gvl = price_chg
            End If
            
            Set rng = Range(output).Offset(0, 1)
            
            'You should also have conditional formatting that will highlight positive change in green and negative change in red.
            For Each Cell In rng
                If Cell.Value > 0 Then
                    Cell.Interior.ColorIndex = 4
                ElseIf Cell.Value < 0 Then
                    Cell.Interior.ColorIndex = 3
                Else
                    Cell.Inerior.ColorIndex = 0
                End If
            Next

            'update mark
            output = Range(output).Offset(1, 0).Address(0, 0)
            mark = ActiveCell.Offset(1, 0).Address(0, 0)
            Debug.Print "New Mark"; mark
            
        Else
            Total = Total + ActiveCell.Offset(0, 5).Value
        End If
        'Next
        ActiveCell.Offset(1, 0).Select
    Loop
    
    Range("P2").Value = git
    Range("Q2").Value = gip
    Range("P3").Value = gdt
    Range("Q3").Value = gdp
    Range("P4").Value = gvt
    Range("Q4").Value = gvl

End Sub

