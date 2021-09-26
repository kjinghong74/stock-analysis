***All Stocks Analysis with VBA***

## **Overview of Project**
Daily data for 12 stocks in year 2017 and 2018 are collected including the date, their open prices and close prices, daily volume, etc. VBA codes in module 2 is the original codes generated to analyze all stocks data in year 2017 and year 2018 and show the total daily volume and return for each stock in each year. In module 3, VBA code is refactored using tickIndex in order to increase the speed of running time.

## **Results**
- Analysis results for year 2017: Overall, the performance of these stocks is doing very well.Except for TERP with 7.2% return loss, all other 11 stocks have positive return. Four of them (DQ, ENPH, FSLR and SEDG) have yearly gain over 100%. DQ has the best return nearly 200%, although DQ's total taily volume is the lowest one in this set of stocks. It doesn't looks like there is correlation between total daily volume and return.
- 
![analysis time for year 2017](https://user-images.githubusercontent.com/90361056/134825137-4c2e2320-53bc-4928-89e5-a0d7763d9e0d.PNG)

- Analysis results for year 2018: The performance of these stocks in year 2018 is not good. Majority of them are negative in yearly return. Only ENPH and RUN show positive return about 80%. The total voume for both ENPH and RUN increased 2- 3 times in year 2018. 

![analysis time for year 2018](https://user-images.githubusercontent.com/90361056/134825148-de6ee2fa-f961-425b-8e97-088a401e1a8e.PNG)

- Put the data in 2017 and 2018 together, DQ has the lowest total daily volume and best gain in 2017.  Its total daily volume increased significantly in 2018 but the retrun showed the worst result by lossing almost 63%. ENPH and RUN are doing well two year in a row, and their total daily volume dramaticly increased too. It looks like the popularity and marketing valaue of ENPH and RUN is increasing. We may recommand these two stocks to Steve's parents.   

## **Summary**
- The advantages of refactoring code: refactoring code can make the code more efficient and run faster. Other advantages include arranging code in a more logic and clear way, which make it easily understood by other viewer. These advantages make the same code easily modified and applied to other purposes.

- The disadvantages of refactoring code: refactoring code takes more time and code writer need invest extra thinking into it. It can generate more variables and makes code have more layers and complicated

-The influence of refactoring on the original VBA script on All Staock Analysis: in this analysis, I refactored the original code by using tickIndex to refer each stock and its related data points. The output parameters are organized in the form of arrays and refered by tickerIndext. In this way, computer need less time to run refactored code compare to original code.   
