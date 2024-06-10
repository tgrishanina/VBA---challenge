#  VBA Project

Create a script that loops through all the stocks for each quarter and outputs the following information:

* The ticker symbol

* Quarterly change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter.

* The percentage change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter.

* The total stock volume of the stock. 

* Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume".

* Make the appropriate adjustments to your VBA script to enable it to run on every worksheet (that is, every quarter) at once.


# Note

* Make sure to use conditional formatting that will highlight positive change in green and negative change in red.


# Other Considerations

* Use the sheet alphabetical_testing.xlsx while developing your code. This dataset is smaller and will allow you to test faster. Your code should run on this file in just a few seconds.

* Make sure that the script acts the same on every sheet. The joy of VBA is that it takes the tediousness out of repetitive tasks with the click of a button.


# Process

1) Define all necessary variables for calculating values and looping through the worksheets.

2) Use a "for" loop to cycle through all of the worksheets.

3) Within that loop, set up another "for" loop to specify what we want to do within each worksheet.

4) We need to comb through the data for each stock symbol and determine the opening quarterly price, closing price, and the total volume throughout the quarter. To do this, we set up another loop to check through the entire dataset.

5) "If" statements: first, we use an "if" statement to check if the current row we're working with is the first occurence of that ticker symbol. If it is, we identify the opening price and set our Opening_Price variable to that value.

6) We then use a new "if" statement to see if the current row is the last occurrence of that ticker symbol. If so, we identify the closing price, calculate the price change using our Opening_Price variable from before, calculate the percent change, and add the volume value to the total volume. If it is not the final occurence of that symbol, we simply add the volume to the total volume before moving on to the next row.

7) Within this statement, we create a summary table displaying the values we found: the quarterly price change, percent change, and total volume for each symbol.

8) We use another "if" statement to add color coding to the change columns: red for negative changes, and green for positive.

9) Use the Min and Max functions to find the greatest increase, decrease, and volume.

10) Use the Match function to find the ticker symbols that correspond to those values.

11) Create a small table displaying the greatest increase, decrease and volume and the corresponding ticker symbols.


# Code Sources

* Received help from Tutor and Learning Assistant on Lines 106- 114 (Match Function). 
* Used Tutor's guidance to google how to loop through worksheets.
* Used Learning Assistant for general help with syntax.
