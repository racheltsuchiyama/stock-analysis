# stock-analysis

In order to make informed decisions for trading stocks, I created a macro in VBA to find the number of trades and returns for each stock. The number of trades per day, or the total daily volume, helps us gauge the accuracy of the stock price. The actual value of the stock converges with the stock price as the frequency of trades increase. The returns represent the percentage increase or decrease of the price at the end of the year compared to the beginning of the year. With these two values, we can wisely choose which stocks to invest in.

## Results of all stocks analysis
### Original macro
First, I created a macro to calculate the total daily volume, starting price, and ending price of each stock using nested for loops. The code first loops through an array of all the stock tickers. For each ticker, the nested loop runs through all the rows of data to determine these values.

![original](https://user-images.githubusercontent.com/83552696/117517055-a9254000-af4f-11eb-8e45-a380d32a1372.png)

The return was then calculated by determining the percent change of the starting price to the ending price. This program was successful in performing the necessary analysis, but did not include any formatting. 

![original_script2017](https://user-images.githubusercontent.com/83552696/117517502-0d94cf00-af51-11eb-94ab-19989c7cdb68.PNG)
![original_script2018](https://user-images.githubusercontent.com/83552696/117517507-0ec5fc00-af51-11eb-8648-496a8fa67e84.PNG)

The information is all readily available, but not in an easily read format.

