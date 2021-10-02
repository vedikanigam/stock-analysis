# Stock-Analysis

## Overview of Project
This project aims to help Steve in knowing how DAQO and other green energy companies are performing on stock market in the years 2017 and 2018 so that his parents can diversifying their investment. Using, a programming language VBA (Visual Basic Application) which is an extension of Microsoft Excel, we will analyze how various green energy stocks performed in the year 2017 and 2018.

### Purpose
The purpose of the analysis is to compare the stocks of all green energy companies in 2017 and 2018 by writing a code in VBA that will help automate the calculations required to perform this analysis. This will greatly reduce the amount of time required to complete the analysis and decrease chances of any errors if it was done only in excel. Furthermore, to apply this code to entire stock market versus a few dozen green energy stocks, we will refactor our original code to reduce the runtime of this code even further.

## Results

### Stock performance in 2017 and 2018
After running our code, we can see that in the year 2017, all the companies had a positive rate of return except company TERP which had a negative rate of return of -7.2%. DQ was the most profitable company with 199.4% rate of return. The other companies that had over 100% rate of return were ENPH, FSLR, and SEDG. See table below. 

![All Stocks Analysis 2017](https://github.com/vedikanigam/stock-analysis/blob/main/Resources/AllStocksAnalysis2017.png)


In the year 2018, most of the companies had a negative rate of return except ENPH and RUN with 81.9% and 84% respectively. DQ was the worst performer with -62.6% rate of return. The other firms that were below -30% rate of return was JKS (-60.5%), SPWR (-44.6%), and FSLR (-39.5%). See table below. 
Seeing DQ’s performance from 2017 to 2018, it could be said that this company’s stocks are volatile. It moved from being the most profitable in 2017 to the most unprofitable in 2018. 

![All Stocks Analysis 2018](https://github.com/vedikanigam/stock-analysis/blob/main/Resources/AllStocksAnalysis2018.png)


### Original Code and its execution time

In the original code, we create a loop that goes from row 2 to row end to calculate total volume traded and percentage change in price for each ticker. Since, this loop is nested within the loop that goes from ticker (0) to ticker (11) the computer goes over all the rows in worksheet 2017 or 2018 each time it performs calculations for individual ticker. Furthermore, since we are printing the output calculations in “All Stocks Analysis “worksheet, it jumps from worksheet 2017 or 2018 to print the outcome in Analysis worksheet and goes back to worksheet 2017 or 2018 to repeat the process for next ticker. While, it was great for Steve to see the results so quickly using our code, this code will take much longer time when applied to entire stock market. 

![All Stocks Analysis Code](https://github.com/vedikanigam/stock-analysis/blob/main/Resources/AllStocksAnalysis_Code.png)

The execution time for this code was 0.84375 seconds for the year 2017 and 0.8203125 seconds for the year 2018. We will refactor our code to reduce runtime so the code can work for large number of stocks. 

![All Stocks Analysis 2017 Runtime](https://github.com/vedikanigam/stock-analysis/blob/main/Resources/AllStocksAnalysis2017_Runtime.png)


![All Stocks Analysis 2017 Runtime](https://github.com/vedikanigam/stock-analysis/blob/main/Resources/AllStocksAnalysis2018_Runtime.png)


### Refactored Code and its execution time
