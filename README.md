# **Stock Data Analysis**

## Overview of Project

### Purpose
The purpose of this analysis is to provide our client with an Excel workbook including an easy-to-run VBA macro able to analyze an entire dataset of stocks. This tool will help him in its financial expertise.\
The analysis was ran using two VBA scripts: the original script produced through the Module #2 and a refactored version of it. We would then have the opportunity to compare the performances of both scripts, and highlights the pros and cons of refactoring a code.

## Results

### Stock Performance Comparison Between 2017 and 2018

These are the screenshots of the stocks performance in 2017 and 2018 obtained with both scripts:

![2017_results](https://user-images.githubusercontent.com/68669675/89343442-e7609c80-d669-11ea-92ba-3c8f5b790035.png)
![2018_results](https://user-images.githubusercontent.com/68669675/89343444-e7f93300-d669-11ea-83bc-77f4430f31e5.png)

In 2017, all stocks except TERP have a positive return. 4 stocks even have over 100% return: DQ, ENPH, FSLR and SEDG. The best performing one was DQ with 199.4% return!\
This was a performing year for the green stocks.\
\
In 2018 mostly all green stock returns went in the red. Only ENPH and RUN continued to perform with respectively 81.9 and 84.0% that year.\
ENPH has the highest total daily volume and was the most traded stock that year.

ENPH is globally the best performing green stock over 2017 and 2018.

### Performance Comparison Between Original and Refactored Scripts
The analysis was ran with the original VBA script obtained through Module #2 and the refactored version of that script.\
The idea is to decrease the processing time of the program by only going through all the dataset one time and retrieve all the information.\
Let's have a look at the code from the original script:
```
'Loop over the tickers array
    For i = 0 To 11
    
        ticker = tickers(i)
        totalVolume = 0
        'Loop over the data
        Worksheets(yearValue).Activate
        For j = 2 To rowEnd
            'totalVolume for the current ticker
            If Cells(j, 1).Value = ticker Then
                totalVolume = totalVolume + Cells(j, 8).Value
            End If
            
            'startingPrice for the current ticker
            If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
                startingPrice = Cells(j, 6).Value
            End If
            
            'endingPrice for the current ticker
            If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then
                endingPrice = Cells(j, 6).Value
            End If
        Next j
        'Output results
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
    Next i
```
Here the script goes through nested loops over all the data with the variable "j" for each ticker with the variable "i". For 12 tickers, the program goes through all the dataset 12 times.

The refactored version uses arrays for the results that are filled along going through all the data rows only one time. Here is the refactored code:
```
'1a) Create a ticker Index
    Dim tickerIndex As Integer

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    '2a) Initialize ticker volumes to zero
    For tickerIndex = 0 To 11
        tickerVolumes(tickerIndex) = 0
    Next tickerIndex
    're-initialize tickerIndex to zero before looping over all rows
    tickerIndex = 0
        
    '2b) loop over all the rows
    For i = 2 To RowCount
         
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            'starting price for the current ticker
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            'ending price for the current ticker
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
                
        End If
        
        '3d) Increase the tickerIndex.
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
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
```

The execution times of both year analysis with the original VBA script are:

![2017](https://user-images.githubusercontent.com/68669675/89342768-e0855a00-d668-11ea-9eb3-99bceb9abbf5.png)
![2018](https://user-images.githubusercontent.com/68669675/89342770-e11df080-d668-11ea-89fe-44738c1c833b.png)

The execution times of both year analysis with the refactored VBA script are:

![VBA_Challenge_2017](https://user-images.githubusercontent.com/68669675/89342735-d2373e00-d668-11ea-9530-011d6fbc9d61.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/68669675/89342736-d2cfd480-d668-11ea-851d-f70840911378.png) 

Refactoring the script increase tremendously its performance, as seen on the screenshots above the execution time gets shorter.

## Summary

1. What are the advantages or disadvantages of refactoring code?
	- pros: makes code faster. Preserves a clean and maintainable architecture in evolving code. Reduces bugs.
	- cons: No additional functionality. Costs development time. Quality dependent on previous developers work.

2. How do these pros and cons apply to refactoring the original VBA script?
	- The execution time has improved and the code is clearer and easier to adjust for future updates.
	- The refactored script does not give ability to analyze any set or the whole set of existing stocks. The functionality is the same as the original VBA script which is to analyse the performances of the set of 12 green stocks.
	- Recoding the refactored script would be needed to give our client the ability to analyze any set of stocks.
