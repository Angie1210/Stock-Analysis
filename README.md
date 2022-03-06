# Stock-Analysis
## Overview of the Project

The purpose of this project is to learn how to use VBA to automate some tasks and to apply them through worksheets. We will be able to analyze different databases using more functions and in a faster way than Excel by itself. In this case we are going to help Steve to choose a company to invest in, by finding the total daily volume of shares to see how often a stock is traded and the yearly return for each stock. We are going to use two macros with different approaches and determine which one is more convenient to use based of the time needed to execute the program. We will be working with Stock Databases from Green Companies.

## Results
### Overview of the project
We created a button for each macro, one named "Analysis per Year" and the other one "Analysis Refactored". In both we created an InputBox to choose with which database or "year" to work with. We wanted to set a timer so we have to create 2 variables for the starting time and the endtime. In the "Analysis per Year" we used an array with the tickers and nested loops that move through the ticker array and loop over all the rows of the database to increase the TotalVolume per ticker and to look for the first and last time each ticker appear in order to set the value for the startingPrice and the ending Price. The execution times are the following.


![All_stock_timer 2017](https://user-images.githubusercontent.com/43548929/156904314-7518e65b-aa11-43b4-8b7d-b40de93273f0.png)
![All_stock_timer2018](https://user-images.githubusercontent.com/43548929/156904313-44a0e72e-3095-4d6c-bdc0-a0b2139fbde7.png)

In the "Analysis Refactored" we created a tickerIndex variable to move across 4 different arrays (tickers, tickerVolumeS, Starting and EndingPrices). We use independent loops, an only one to move across all the rows, at the end we increased the tickerIndex variable, to move to next tickerIndex. The execution times are shorter compared with the other approach, we achieved to reduce the execution time by 70%, as we can see in the next images. 

![VBA_Challenge_2017](https://user-images.githubusercontent.com/43548929/156904659-1bd1a6d0-294f-4a81-8dac-2c5636821ad2.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/43548929/156904667-ea937ab2-1aff-4b60-ad8d-a49840f99d95.png)



We used the same formatting for the output, so it looks the same for both macros
## Summary
VBA is a great tool to reduce repetitive tasks and to decreae the chance of errors.
By refactoring the code,  Its a significant difference if we want to use our code in bigger databases. 

the new code is more efficent and it doesn't have to run to 2 loops at the same time, so it uses less memory.
