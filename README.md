# Stock Analysis
## Overview of Project
### Purpose and Background
The purpose of this project was to take code written in previous modules that analyzes stock data for 12 different companies, and refactor it so that it can scale with substantial increases in data.

The worksheet contains 2017 and 2018 data for 12 different companies. The companies are represented as stock tickers, and the data provide is the open, close, high, and low, as well as the volume over the span of each year. The goal is to analyze the total volume and return on investment for each company. 

## Results
### Analysis
I started the analysis by ensuring that the code for the input box, ticker array, timers, and all formatting were in place beforehand. I compared the script to the first allstocksanalysis to make sure that it was correct. I then wrote code in order for each line in the instructions, and began running the code line by line to make sure there were no bugs. Example of refactored code:

    '1a) Create a ticker Index
        'using Dim to assign variable 0 to tickerIndex as an Integer since it's a whole number
    Dim tickerIndex As Integer

    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    ' If the next row’s ticker doesn’t match, increase the tickerIndex.
    For i = 0 To 11
        tickerVolumes(i) = 0
    Next i
    ''2b) Loop over all the rows in the spreadsheet.
    For j = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(j - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         If Cells(j + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
         End If

            '3d Increase the tickerIndex.
             If Cells(j + 1, 1).Value <> tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
            End If
    
    Next j
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.

     For i = 0 To 11
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
    
   Results after running the code for both 2017 and 2018
<p align="left">
  <img src="https://github.com/armyofkittens/stock-analysis/blob/main/Resources/All%20Stocks%20Analysis%202017.png" width="700"/>
</p>

<p align="left">
  <img src="https://github.com/armyofkittens/stock-analysis/blob/main/Resources/All%20Stocks%20Analysis%202018.png" width="700"/>
</p>

## Summary
### Advantages and Disadvantages of Refactoring Code
One of the things I noticed about refactoring the code is that it makes you think twice about every line written, and if it's the most efficient way of performing that action. Consequently it also lead to me figuring out new ways to perform the same action. This can also lead to some disadvantages as even though you have discovered a new way of performing the same action, it may add additional latency to the subroutine. 

### Conclusion on Refactoring Code
There were huge advantages to refactoring the code. My original script ran in .8 seconds where as my refactored code runs it almost 8 times faster. What I believe caused this is that the calls to the various arrays were done by each individual for loop, rather than within nested for loops. Refactoring the code and lowering the latency gives the subroutine more headspace in the future to analyze even larger pools of stock data amongst a broader range of companies. 

Before Refactor 2017

<p align="left">
  <img src="https://github.com/armyofkittens/stock-analysis/blob/main/Resources/elapsed%202017.png" width="350"/>
</p>

After Refactor 2017

<p align="left">
  <img src="https://github.com/armyofkittens/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png" width="350"/>
</p>

Before Refactor 2018

<p align="left">
  <img src="https://github.com/armyofkittens/stock-analysis/blob/main/Resources/elapsed%202018.png" width="350"/>
</p>

After Refactor 2018

<p align="left">
  <img src="https://github.com/armyofkittens/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png" width="350"/>
</p>
