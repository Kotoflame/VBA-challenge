Attribute VB_Name = "Module1"
Sub StockLoop():
    
    'Written by Ryan C for UMN Data Analytics Short Course 2023
    '4/08/2023
    'Code takes ~30 seconds total on a Ryzen 5 3600, so most modern processors should handle it fine
    'All formatting is done automatically
    
    
    
    'Initializing used variables
    Dim Endrow As Long
    Dim TickerCount, Yr As Integer
    Dim Startprice, Endprice, YearChange, StockVolume As Double
    Dim GreatestIncrease, GreatestDecrease, GreatestVolume As Double
    Dim CurrentTicker As String
    Dim Location As String
    Dim c As Range
    


    
    
    'Looping over all of the sheets (in this case named per years)
    For Yr = 2018 To 2020
        Worksheets(CStr(Yr)).Activate
    
        'Adjusting column width per my sanity
        For i = 1 To 17
            Columns(i).ColumnWidth = 15.3
        Next
        Columns("O").ColumnWidth = 21
        Columns("M").ColumnWidth = 2.5
        Columns("N").ColumnWidth = 2.5
        Columns("L").ColumnWidth = 17.5
        
        'Getting the last filled row in the WS
        Endrow = Worksheets(ActiveSheet.Name).UsedRange.Rows.Count
        

        'Setting up the column names for the results
        Cells(1, "I").Value = "Ticker"
        Cells(1, "J").Value = "Yearly Change"
        Cells(1, "K").Value = "Percent Change"
        Cells(1, "L").Value = "Total Stock Volume"
        Range("K:K").NumberFormat = "0.00%"    'formatting K as a percent.. this does cause problems later
        Range("J:J").NumberFormat = "0.00"    'formatting J as a 0.00
        Range("Q2:Q3").NumberFormat = "0.00%"    'formatting best and worst % change
        Range("Q4").NumberFormat = "0.00E+00"     'formating the max total volume cell
        'Starting Ticker counter at 1
        TickerCount = 1
        
        For i = 2 To Endrow 'looping of all rows with data
            
            If i = 2 Then  'on the loop initiation, I grab the first tickers initial price manually. I could probably avoid this, but it works
                Startprice = Cells(i, "C").Value
            End If
            
            
            CurrentTicker = Cells(i, "A").Value 'Current ticker for row i
            StockVolume = StockVolume + Cells(i, "G").Value 'running total of stock volume
            
            If CurrentTicker <> Cells(i + 1, "A") Then 'if the next ticker is different
                Endprice = Cells(i, "F").Value 'Grab the year end close value and do the math for the current ticker
                YearChange = Endprice - Startprice
                Cells(TickerCount + 1, "I").Value = CurrentTicker 'fill in the results table for current ticker
                Cells(TickerCount + 1, "J").Value = YearChange
                Cells(TickerCount + 1, "K").Value = YearChange / Startprice
                Cells(TickerCount + 1, "L").Value = StockVolume
                If YearChange >= 0 Then 'Red/Green conditional formatting for yearly change
                Cells(TickerCount + 1, "J").Interior.Color = RGB(10, 200, 10)
                Else
                Cells(TickerCount + 1, "J").Interior.Color = RGB(255, 20, 20)
                End If
                
                
                TickerCount = TickerCount + 1 'The next ticker is different, add to my ticker count tracker
                Startprice = Cells(i + 1, "C").Value 'collect the next tickers start price
                StockVolume = 0 ' reset my running Stock Volume tracker
            End If
              
        
        Next
        
        'Set up the headers for my final results table
        Cells(2, "O").Value = "Greatest % Increase"
        Cells(3, "O").Value = "Greatest % Decrease"
        Cells(4, "O").Value = "Greatest Total Volume"
        Cells(1, "P").Value = "Ticker"
        Cells(1, "Q").Value = "Value"
        
        
        GreatestIncrease = Application.WorksheetFunction.Max(Range("K:K")) 'Get the Max of the %change row
            With Worksheets(ActiveSheet.Name).Range("K:K") 'my preffered .find method using with to get a cell address
                Set c = .Find(GreatestIncrease * 100) 'This *100 is why the percent formatting is really annoying
                Location = c.Address
            End With
            CurrentTicker = Range(Location).Offset(0, -2).Value 'offset -2 to get the relative tracker for the found max address
            Cells(2, "P").Value = CurrentTicker
            Cells(2, "Q").Value = GreatestIncrease
        
        'Repeat above, minimum command
        GreatestDecrease = Application.WorksheetFunction.Min(Range("K:K"))
            With Worksheets(ActiveSheet.Name).Range("K:K")
                Set c = .Find(GreatestDecrease * 100)
                Location = c.Address
            End With
            CurrentTicker = Range(Location).Offset(0, -2).Value
            Cells(3, "P").Value = CurrentTicker
            Cells(3, "Q").Value = GreatestDecrease
        
        'Repeat above, volume traded amounts
        GreatestVolume = Application.WorksheetFunction.Max(Range("L:L"))
            With Worksheets(ActiveSheet.Name).Range("L:L")
                Set c = .Find(GreatestVolume)
                Location = c.Address
            End With
            CurrentTicker = Range(Location).Offset(0, -3).Value '-3 offset for extra column
            Cells(4, "P").Value = CurrentTicker
            Cells(4, "Q").Value = GreatestVolume
        
    
    Next Yr  'Look to the next year (sheet)

End Sub



