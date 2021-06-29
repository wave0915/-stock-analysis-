# **Stock-Analysis**
## Overview

Performing analysis by using VBA to perform financial analysis on [green_stocks](green_stocks.xlsm) data to help Steve's parents to decide whether they should invest all the money in DQ stocks or not.  We are going to compare the green energy stock data in 2017 and in 2018  by using macro in excel.

## Results

### -2017 vs. 2018 in stock performance
To compare the performance of 2017 and 2018, first use conditional statement in for loop to filter the starting and ending prices of current ticker(DQ) by using the code below in excel macro.
       

    For i = 2 To RowCount
    
        'Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        'Find starting price for current ticker
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        'Find ending price for current ticker
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            tickerIndex = tickerIndex + 1
        End If 
    Next i
    
Then, use for for loop to find the outcome of DQ in percentage by using code below.

      For i = 0 To 11
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
      Next i
      Range("B4:B15").NumberFormat = "#,##0"
      Range("C4:C15").NumberFormat = "0.0%"
      
DQ stock was showing increase by 199.4% in 2017 return and decrese by -62.6% in 2018 return. Also, total daily volumne in 2017 was 35,796,200 and in 2018 was 107,873,900, which show the volumne has been increased as shown the charts below.

- All Stocks(2017)
![2017](resources/VBA_Challenge_2017.png)

- All Stocks(2018)
![2018](resources/VBA_Challenge_2018.png)




