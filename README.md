# VBA of Wall Street

## Overview of Project
### Purpose
Steve's parents are looking to invest into green energy and he wishes to provide them with accurate information for a sound investment decision. The purpose of this project is to find annual volumes and returns for select stocks as well as refactor the initial VBA code to run more efficiently in case Steve wishes to look at a larger amount of data. To that end, Steve has compiled start and end prices and daily trading volumes for twelve different stocks over the course of 2017 and 2018.

## Results
### Stock Performance
According to the data, 2017 seemed to be a good year for investing in green energy as every stock except for TERP had a positive return. In fact, the top two returns were both around 200% and included the stock Steve's parents had selected (DQ). By 2018, however, only two stocks maintained their positive return - ENPH and RUN - and even so, ENPH's returns declined from just over 136% to just under 98%. RUN was the only stock with significantly improved performance, with returns increasing from 9.5% in 2017 to 85.2% in 2018. DQ's performance was actually the worst in 2018 with a -61.1% return which, given that it had the highest performance in 2017, also meant that it had the greatest decrease in returns percentage-wise.

The below images show the data for 2017 and 2018.

![This is an image](https://github.com/EricaEidelman/stock-analysis/blob/main/2017%20Data.png)

![This is an image](https://github.com/EricaEidelman/stock-analysis/blob/main/2018%20Data.png)

In general, volumes were higher in 2018 although 2017 saw the two highest volumes for an individual stock for SPWR and FSLR, respectively. Incidentally, these two stocks saw a drop in volume and were among the lowest performing stocks in  2018, which may indicate that an adverse event precipitated a drop in value and selling off of the two stocks. While DQ saw an increase in volume from 2017 to 2018, its significantly reduced returns may also indicate an upcoming volume reduction.

### Script Execution Times
When comparing the original and the refactored code, the refactored code takes an average of about 0.175 seconds (about 0.15 for 2017 and 0.20 for 2018) compared to the original's 0.82, or about 4.68 times quicker. Images of run times for the refactored code are below.

![This is an image](https://github.com/EricaEidelman/stock-analysis/blob/main/VBA_Challenge_2017.png)

![This is an image](https://github.com/EricaEidelman/stock-analysis/blob/main/VBA_Challenge_2018.png)

Both the original and refactored code create header rows for the data as well as an array of the stock tickers to match them with their volume and return data. However, the original code loops through all the rows of data for every single ticker when calculating annual volume and returns. The refactored code recognizes that the stock data is arranged by ticker symbol and so once the last row for a given ticker is reached, there will be no more data for that ticker. Therefore, the refactored code loops through all the data rows only once, changing the ticker symbol for its calculations once the last row for the previous ticker is reached. 

Below is the snippet from the original code calculating volumes and returns. The variable i refers to the position of each ticker symbol in the array of all symbols and the variable j refers to the row number in the data. With just over 3,000 rows of data and 12 ticker symbols, the computer has to run the if-then statements over 36,000 times. The starting and ending prices are noted for the first and last rows, respectively, of each ticker symbol.

```
    For i = 0 To 11
        ticker = tickers(i)
        totalVolume = 0
  
    Worksheets(yearValue).Activate
    
        For j = 2 To RowCount
        
            If Cells(j, 1).Value = ticker Then
                totalVolume = totalVolume + Cells(j, 8).Value
            End If
        
            If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
                startingPrice = Cells(j, 3).Value
            End If
            
            If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then
                endingPrice = Cells(j, 6).Value
            End If
        
        Next j
        
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
    
    Next i
```

And following is the snippet from the refactored code. The first for loop sets the initial volume of each ticker symbol at 0. What follows is a command to go through the lines of data increasing the volume of a ticker symbol until its last row is reached, when the ticker index is increased and the computer moves on to the next stock. In this code, the computer loops through all the data only once, meaning it has to run the if-then statements just over 3,000 times, unlike the 36,000 times of the original code. Starting and ending prices are found as above.

```
    Dim tickerIndex As Integer
    tickerIndex = 0

    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    For i = 0 To 11
        tickerVolumes(i) = 0
    Next i
        
    For i = 2 To RowCount
    
        If Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        End If
        
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 3).Value
            
         End If
 
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
         End If

         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
            
         End If
    
    Next i
    ```
    
## Summary
- What are the advantages of disadvantages of refactoring code?

- How do these pros and cons apply to refactoring the original VBA script?
