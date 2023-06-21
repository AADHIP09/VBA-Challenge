# VBA-Challenge

References: 
The below sections of code were based on information from : https://www.wallstreetmojo.com/vba-max/
 
            TableRange = ws.Range("K1:K" & TableRow)
            TotalVolumeRange = ws.Range("L1:L" & TableRow)
            GreatestIncrease = WorksheetFunction.Max(TableRange)
            GreatestDecrease = WorksheetFunction.Min(TableRange)
            GreatestVolume = WorksheetFunction.Max(TotalVolumeRange)


The below section of code was based on information here: https://www.mrexcel.com/board/threads/help-with-avoiding-division-by-zero-error-in-vba.783862/


             If (openPrice = 0) Then
             percentChange = 0
        
            'calculating percentage change in comparison to opening price
             Else
             percentChange = (priceChange / openPrice)

The formatting in cells to show percentage values in various parts of code is based on this: https://stackoverflow.com/questions/20648149/what-are-numberformat-options-in-excel-vba
Specific lines of code given below: 

            ws.Range("K" & Row).Value = percentChange
            ws.Range("K" & Row).NumberFormat = "0.00%"

            ws.Range("Q2").Value = GreatestIncrease
            ws.Range("Q2").NumberFormat = "0.00%"
            ws.Range("Q3").Value = GreatestDecrease
            ws.Range("Q3").NumberFormat = "0.00%"

