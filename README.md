# stock-analysis

In order to make informed decisions for trading stocks, I created a macro in VBA to find the number of trades and returns for each stock. The number of trades per day, or the total daily volume, helps us gauge the accuracy of the stock price. The actual value of the stock converges with the stock price as the frequency of trades increases. The returns represent the percentage increase or decrease of the price at the end of the year compared to the beginning of the year. With these two values, we can wisely choose which stocks to invest in.

## Results of all stocks analysis
### Original macro
First, I created a macro to calculate the total daily volume, starting price, and ending price of each stock using nested for loops. The code first loops through an array of all the stock tickers. For each ticker, the nested loop runs through all the rows of data to determine these values.

![original](https://user-images.githubusercontent.com/83552696/117517055-a9254000-af4f-11eb-8e45-a380d32a1372.png)

The return was then calculated by determining the percent change of the starting price to the ending price. This program was successful in performing the necessary analysis, but did not include any formatting. 

![original_script2017](https://user-images.githubusercontent.com/83552696/117517502-0d94cf00-af51-11eb-94ab-19989c7cdb68.PNG)
![original_script2018](https://user-images.githubusercontent.com/83552696/117517507-0ec5fc00-af51-11eb-8648-496a8fa67e84.PNG)

The information is all readily available, but not in an easily read format. I created a different macro to convert the returns to percentages and added color for quick extraction of a percent increase or decrease. Instead of having to run several macros to make this analysis more user friendly, I refractored the script.

### Refactored macro
Since nested loops can get complicated, I utilized a new variable called the tickerIndex. Instead of iterating through all the tickers and finding the total volumes and returns, I made arrays to hold the forementioned values for each ticker. The new code uses the tickerIndex to find the trade volume and returns for each stock, but differs from the original code by storing these values in an array. The tickerIndex was initially set to zero, but increases by one at the end of the loop.

![refactored](https://user-images.githubusercontent.com/83552696/117555300-e1418700-b012-11eb-900a-a8fb9cc7731d.png)

Then, he refactored code formats the values and adds color to represent the positive or negative values of the return. While this code performs more functions than the original code, it has a longer run time as displaye below.

![VBA_Challenge_2017](https://user-images.githubusercontent.com/83552696/117525311-72165500-af76-11eb-96b8-5bbb5020aae2.PNG)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/83552696/117525314-7478af00-af76-11eb-9136-c829473d3100.PNG)

## Summary
The original code has a slight advantage in run time in comparison to the refactored code. However, this advantage could be void due to the increased functions of the refactored code. The refactored code adds commas to the daily total volume and converts the return values to percentages. It also changes the cell color to green for a positive percentage increase and to red for a negative percentage decrease. These factors make the slight increase in run time insignificant for the refacotred code. While the original code runs faster than the refactored code, the addition functions of the refactored code make it more useful to my analysis.

