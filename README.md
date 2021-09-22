# Stock Analysis
Analysis of green energy stock.

## Project Overview

This analysis was put together for Steve and his Parents. Steve's parents want to invest in green energy stick. They have chosen DAQO new energy corp (DQ). DQ makes silicon parts for solar panels. Steve wants to make sure his parents have chosen a profitable green stock to invest in. To answer this question, this analysis compares DQ stock data with other green energy organizations. VBA was used to find total daily volume and and yearly return for each stock. 

  total volume = total # of shares traded throughout the day -measure of how actively a stock is being traded 
  yearly return = % difference in price from start of year to end of year
  
 
## Results 

VBA was used to find total daily volume and and yearly return for each stock. *VBA_Challenge.xlsm* contains green stock data for 2017 and 2018. First, an array called *tickers* of green organization names was greated. This 12 strings of names are refered to as tickers and allow us to loop through the stock data and perform analyses on one organisation at a time. Three arrays with a length of 12 (there are 12 green stock companies to be analysed) were created to hold the output from this analysis: *tickerVolumes*, *tickerStartingPrices*, and *tickerEndingPrices*. 

Next, a nested for loop was implemented. The outer loop takes a organisation name from *tickers* based on an index. The inner loop uses this organisation's name from *ticker* to loop throught rows of stock data. If else statements allow VBO to:
  1. match the ticker name with an organisation's name in column A
  2. sum all volumes 
  3. find the stock price for the start of the year
  4. find the stock price for the end of the year

Once the inner loop has worked through all the tockers (all 12 green companies), the outer loop puts the data now entered in to the output arays a new excel sheet: 
  - Company Names = *tickers*
  - Volume = *tickerVolumes*
  - Return = *tickerEndingPrices* / *tickerStartingPrices* - 1

This VBA code was written twice. The sencond time was an attempt at making the code run more quicky (the code was refactored). The Timer fucntion was used to find the toal amount of time it took to perform each version of the analysis. Below is a screen shot of the message box produced by VBA telling us the run time. 



xxx method was faster. 

## Summary

There is a detailed statement on the advantages and disadvantages of refactoring code in general (3 pt).
There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script (3 pt).


