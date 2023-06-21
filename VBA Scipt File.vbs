Sub StockAnalysis():

'declaring the variables

Dim i As Double
Dim Total As Double
Dim Name As String
Dim Row As Long
Dim ws As Worksheet
Dim openPrice As Double
Dim closePrice As Double
Dim priceChange As Double
Dim percentChange As Double
Dim TableRow As Double
Dim GreatestIncrease As Double
Dim GreatestDecrease As Double
Dim GreatestVolume As Double
Dim TableRange As Variant
Dim TotalVolumeRange As Variant

For Each ws In Worksheets

'creating new table headers
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percentage Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(2, 15).Value = "Greatest Percentage Increase"
ws.Cells(3, 15).Value = "Greatest Percentage Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"



'assigning starting values to Variables
openPrice = ws.Cells(2, 3).Value
closePrice = 0
Total = 0
Row = 2

      
'looping through the rows

    For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    
   'finding the point in the sheet where the next value is not the same as the current value
   
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
  
       'extracting the ticker name
        Name = ws.Cells(i, 1).Value
        
       'computing the total for each ticker
        Total = Total + ws.Cells(i, 7).Value
       
       'extracting closing price at the end of the year
        closePrice = ws.Cells(i, 6).Value
        
       'calculating price change from end of the year to start of the year
        priceChange = (closePrice - openPrice)
       
            'solving the divisibility by zero issue
             If (openPrice = 0) Then
             percentChange = 0
        
            'calculating percentage change in comparison to opening price
             Else
             percentChange = (priceChange / openPrice)
                  
              End If
        
        
      'displaying calculated values in the sheet
        ws.Range("I" & Row).Value = Name
        ws.Range("J" & Row).Value = priceChange
        ws.Range("K" & Row).Value = percentChange
        ws.Range("K" & Row).NumberFormat = "0.00%"
        ws.Range("L" & Row).Value = Total
        
       'jumping to next row for next ticker
        Row = Row + 1
        'resetting the starting Total value to 0 for next ticker
        Total = 0
        
        'obtaining the opening price for each ticker
        openPrice = ws.Cells(i + 1, 3).Value
       
        
        
        Else
    
        Total = Total + ws.Cells(i, 7).Value
        
         
        
       End If
  
  'closing first loop
  Next i
  
  
  'adding conditional formatting using a new loop
        TableRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
  
  'initating new loop for Summary Table
  For i = 2 To TableRow
    
    If ws.Cells(i, 10).Value > 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 4
        
    Else
        ws.Cells(i, 10).Interior.ColorIndex = 3
     
    End If
    
    'creating range for row K and row L to make calculation of maxmimum and minimum values easier
    
    TableRange = ws.Range("K1:K" & TableRow)
    
    TotalVolumeRange = ws.Range("L1:L" & TableRow)
      
    'using Max function in Excel to find out the max value from Column K
    GreatestIncrease = WorksheetFunction.Max(TableRange)
    
    'using If statement to identify the respective Ticker for the Maximum value in Column K & displaying it in Column P
    If ws.Cells(i, 11).Value = GreatestIncrease Then
    ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
    
    End If

    
    'using Min function in Excel to find out the Min value from Column K
    GreatestDecrease = WorksheetFunction.Min(TableRange)
       
    'using If statement to identify the respective Ticker for the Minimum value in Column K & displaying it in column P
    If ws.Cells(i, 11).Value = GreatestDecrease Then
    ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
    
    End If
    
    'using Max function in Excel to find out the max value from Column L
    GreatestVolume = WorksheetFunction.Max(TotalVolumeRange)
    
    'using If statement to identify the respective ticker for the Maximum value in column L & displaying it in column P
    If ws.Cells(i, 12).Value = GreatestVolume Then
    ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
    
    End If
    
    
    
    'Displaying all values computed above
    
    ws.Range("Q2").Value = GreatestIncrease
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("Q3").Value = GreatestDecrease
    ws.Range("Q3").NumberFormat = "0.00%"
    ws.Range("Q4").Value = GreatestVolume
     
    

  Next i
      
      
Next ws

 


End Sub





