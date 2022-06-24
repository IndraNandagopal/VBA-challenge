# The VBA of Wall Street

## Background

This project is to analyse the stock market data for 3 years (2018, 2019 and 2020) using VBA scripts.

### Files Used

* [Test Data](\alphabetical_testing.xlsm) - Used this while developing the VBA script.

* [Stock Data](\Multiple_year_stock_data.xlsm) - Ran the script on this data to generate the final report.

### Solution

The goal is to create a VBA script that loops through all the stocks for three years and output the following information for each year:

  * The ticker symbol.

  * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The total stock volume of the stock.

**Note:** Made sure to use conditional formatting that will highlight positive change in green and negative change in red.

Added functionality to the script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume".

Made the appropriate adjustments to the VBA script to allow it to run on every worksheet (that is, every year) just by running the VBA script once.


## Output

Uploaded the following to GitHub:

  * A screen shot of results for each year on the multi-year stock data. These screen shots are available in "Resources" folder.

### 2018:
![2018 Stock Data analysis](Resources/hard_solution_2018.PNG)

### 2019:
![2019 Stock Data analysis](Resources/hard_solution_2019.PNG)

### 2020:
![2020 Stock Data analysis](Resources/hard_solution_2020.PNG)

  * VBA script as a separate file. 
    The name of the VBA script is "stockData.vbs"



