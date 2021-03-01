# Stock Analysis

### Brian Gerrard's Analysis on Stock Data 
---
## Overview of Project
For this project I analyzed stock information from 2017 and 2018 for Steve’s parents to help them determine whether or not the stocks are worth investing in. The data consists of 12 different stocks, their ticker value, stock issue date, opening/closing price, highest/lowest stock price and volume.  

I’ll also be focusing on the below data points:
The total daily volume and yearly return for each stock:

-	**Total Daily Volume** = total number of shares traded throughout the day, this measures how actively a stock is traded on a daily basis

-	**Yearly Return** = percentage difference in price from the beginning of the year to the end of the year 

# Results

Below are the steps I took to refactor the Module 2 solution code:


    '1a) Create a ticker Index
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
    Next i
    
     
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
         
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
            
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
         
        End If
            
            '3d Increase the tickerIndex.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
            End If
            
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub

# Execution Time Results:
You can see below that from refactoring the code there is a significant decrease in run time compared to the original script from Module 2. The 2017 refactored script ran ~6 hundredths of a second faster than the original script while the 2018 refactored script ran ~7 hundredths of a second faster than the original script. 


## Refactored Script Execution Times:
![VBA_Challenge_2017](https://user-images.githubusercontent.com/75700317/109439390-1a022600-79fc-11eb-8223-42a4fbfe4dcb.png) ![VBA_Challenge_2018](https://user-images.githubusercontent.com/75700317/109439400-238b8e00-79fc-11eb-9c1c-4955038fedfe.png)



## Original Script Execution Times:
![old-2017](https://user-images.githubusercontent.com/75700317/109439404-2ab29c00-79fc-11eb-9681-e2535d903231.png) ![old-2018](https://user-images.githubusercontent.com/75700317/109439407-2e462300-79fc-11eb-91ab-55b65f89bdb9.png)




# Summary:

**Advantages and Disadvantages of refactoring code:**

The less code you have, the easier it is to understand and modify the code. This also makes it easier to spot bugs. You won’t have to go back and try to understand the past code to fix bugs. 

**Pros and Cons of applying refactoring to original VBA script:**

Refactoring makes code more easily understandable and easy to extend in the future as code becomes more complex.  

Refactoring code eliminates the amount of code you will have to work with, and better understand the system 

Though eliminating duplication makes modification easeier, one con is that it forces design patterns into your code which…
