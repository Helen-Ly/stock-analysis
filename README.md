# Challenge

### Objective

The objective of this challenge was to analyze data from specific stocks between 2017 and 2018. We were tasked to use VBA to code a subroutine that would analyze the the annual total volume, and return rate of each stock.

### Method

With our initial code, we created an array for all the stock names as ticker values. We had a line of code that would count how many rows were within each data set. We created a nested loop where the first one looped through the ticker values and the nested loop would loop through all the data set rows to collect the starting and ending price of each ticker and count and add all the volume for each individual tickers.

After collecting the data, we output all the data in a new sheet and formatted the table to highlight the cell in green if the return rate was positive and red if negative.

### Limitations

Though the code ran as we needed, a limitation was how the code was written. The nested loop we created loops through 3013 rows 12 times. It slows down the process of obtaining the information and if the data set happens to be bigger, it will take longer to compile all this information. Not only the time it takes is long, but if someone was to try to understand the code, even with comments, it may take a while to understand what the code is trying to do.

### Solution

What we want to do is refactor our code, which means cleaning it up and making it more efficient. In order to do that, we removed the nested loop and created 4 individual loops. We also declared our variables and set tickerIndex to 0 before our loops.

1. Since we have the tickers array created, the first loop initialized an array for the ticker volume, starting and ending price, and set their values to 0.
2. The second looped through rows 2 to 3013. Within this loop, we had an if statement to add 1 to tickerIndex if the next row did not match the current row. The next if statements collected the starting price, ending price and ticker volume for the current ticker value, which was indicated by our tickerIndex.
3. Our third loop output all the data stored in each variable into a table in our new sheet labeled "All Stocks Analysis". We also formatted the cells within this loop, such as changing the table header to bold.
4. Our last loop added conditional formatting to our return rate column for the table. It states that if the return rate is positive, the cell will be highlighted green and red if it is negative.

### Conclusion

Breaking the nested loop apart allows us to loop through 3013 rows only once to gather our data. As a result, it cuts down the amount of time the code needs to run. Not only that, in addition to adding comments, the code is easier to read, which means it is easier to debug.

After refactoring our code, we can turn our attention to the data. In 2017, our 12 stocks were all showing green for their return rate except for one. However, when you analyze the same stocks in 2018, most of their return rates were in the negatives, with an exception of two that were showing positive. This only shows a glimpse of how these stocks were doing in 2017 and 2018, but more information and furhter analysis is needed to determine which stocks are worth investing in and which are not.
