# Stock-analysis

## Overview of the Project

Steve, our Finance graduate, noticed that ticker: DQ yearly returns were negative and needed a macro built that could run through all the stock data, populate stock trade volume, stock yearly returns, and use that output as a means to advise his parents to consider other stocks. Throughout the project, we defined the stock tickers array, created a loop, a nested loop, conditionals to get sum our volume and gather starting and ending prices, formatted the worksheet to display the information neatly, and conditioned the output to color code positive and negative returns. After several enhacements to our code, which included reducing lines of code by creating a loop to initialize the ticker volume array to zero, and removing the nested loop, and including a timer to measure our execution, we saw a considerably improvement in how quickly the script executed.

## Results

### Stocks: Comparison of Results

The results from the data collected in years 2017 and 2018 are significantly different. In 2017, eleven of the twelve stocks saw positive returns compared to 2018 where only two stocks experienced positive returns. Ticker: ENPH and RUN saw positive results consecutively. 

![2017 Stock Returns](https://github.com/carolinaroca007/Stock-analysis/blob/master/README%20files/All_Stocks_2017.png)

![2018 Stock Returns](https://github.com/carolinaroca007/Stock-analysis/blob/master/README%20files/All_Stocks_2018.png)

### Stocks: Execution Times

Prior to refactoring the subroutine allStocksAnalysis to be more efficient, the script execution time to get all stocks' volumes and returns for 2017 and 2018 were 0.75 and 0.73 second, respectively. I refactored the code by removing the nested loop that would have otherwised run my conditional 11 times and ran the code from A2 to the last row for each ticker. The refactored subroutine now only runs the tickerIndex 


![2017 Script Execution](https://github.com/carolinaroca007/Stock-analysis/blob/master/Resources/VBA_Challenge_2017.png)

![2018 Script Execution](https://github.com/carolinaroca007/Stock-analysis/blob/master/Resources/VBA_Challenge_2018.png)

## Summary

In summary, 
