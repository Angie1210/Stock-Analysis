# Stock-Analysis
## Overview of the Project

The purpose of this project is to learn how to use VBA to automate some tasks and to apply them through worksheets. We will be able to analyze different databases using more functions and in a faster way than Excel by itself. In this case we are going to help Steve to choose a company to invest in, by finding the total daily volume of shares to see how often a stock is traded and the yearly return for each stock. We are going to use two macros with different approaches and determine which one is more convenient to use based of the time needed to execute the program. We will be working with Stock Databases from Green Companies.

## Results
### Overview of the project
We created a button for each macro, one named "Analysis per Year" and the other one "Analysis Refactored". In both we created an InputBox to choose with which database or "year" to work with. We set a timer so we have to create 2 variables for the starting time and the ending time. In the "Analysis per Year" we use an array with the tickers and nested loops that move through the ticker array and loop over all the rows of the database to increase the TotalVolume per ticker and to look for the first and last time each ticker appear in order to set the value for the startingPrice and the ending Price. The execution times are the following.

![All_stock_timer 2017](https://user-images.githubusercontent.com/43548929/156904314-7518e65b-aa11-43b4-8b7d-b40de93273f0.png)
![All_stock_timer2018](https://user-images.githubusercontent.com/43548929/156904313-44a0e72e-3095-4d6c-bdc0-a0b2139fbde7.png)

In the "Analysis Refactored" we created a tickerIndex variable to move across 4 different arrays (tickers, tickerVolumeS, Starting and EndingPrices). We use independent loops, only one to move across all the rows, at the end we increased the tickerIndex variable to start the next loop with the next tickerIndex value. The execution times are shorter compared with the other approach, we achieve to reduce the execution time by 70%, as we can see in the next images. 

![VBA_Challenge_2017](https://user-images.githubusercontent.com/43548929/156904659-1bd1a6d0-294f-4a81-8dac-2c5636821ad2.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/43548929/156904667-ea937ab2-1aff-4b60-ad8d-a49840f99d95.png)

We get the same results per year with both macros. As we can see, in 2017 the best performing stocks were DQ,ENPH, FSLR and SEDG with returns over 100%, but 2018 was a terrible year for the stock market in general and the only ones with positive return were ENPH and RUN,  so based on only these 2 years, we can recommend Steve to invest in **ENPH**, because it also was the most traded.


![StockAnalysis_2017](https://user-images.githubusercontent.com/43548929/156906565-8eaa66b4-d2e5-4cc7-af3a-f85c93dad541.png)
![StockAnalysis_2018](https://user-images.githubusercontent.com/43548929/156906535-9585a6e1-18c3-4569-b65e-a9f3e0faffa0.png)


## Summary
### Refactoring advantages and disadvantages. 
The main advantages of refactoring are that you can make your code more efficient, by simplifying the process and removing redundances.  It also uses less memory by changing the structure, while keeping the same functionalities. The code can be more organized and be easier to understand. It gives scalability and help to reduce the execution times.

The disadvantages are that you can introduce errors or bugs to the code and it's also time consumming.

### Refactoring Stock Analysis

The advantage is that we reduced the execution times by removing the nested loop, so the code will be more efficient with bigger databases. Although the refactored code is faster, it was not easier to understand for me.
