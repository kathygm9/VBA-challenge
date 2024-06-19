VBA Stock Analysis Challenge
============================

Overview
--------
This VBA script automates the analysis of stock data across different quarters.
It calculates essential financial metrics for each ticker symbol and can be
extended to handle multiple data sets across an Excel workbook.

Instructions
------------
# Basic Solution
## Objective
- Create a script that loops through all the stocks for each quarter and outputs
  the following information:
  - Ticker Symbol: Identify each stock.
  - Quarterly Change: Calculate the difference between the opening price at the
    beginning of a quarter and the closing price at the end of that quarter.
  - Percentage Change: Determine the percentage change from the opening price
    at the beginning of a quarter to the closing price at the end of that quarter.
  - Total Stock Volume: Sum the total traded volume of each stock for the quarter.

# Moderate Solution
## Enhancement
- Extend the basic script to also find and return:
  - Greatest % Increase: The stock with the highest percentage increase over the quarter.
  - Greatest % Decrease: The stock with the highest percentage decrease over the quarter.
  - Greatest Total Volume: The stock with the largest total trading volume over the quarter.

# Hard Solution
## Further Extension
- Adjust the script to automatically apply the analysis across all worksheets in an
  Excel workbook, effectively processing each quarter in one run.

