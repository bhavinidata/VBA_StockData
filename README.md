# The VBA of Wall Street

## Task

Use VBA scripting to analyze real stock market data and demonstate mastery on advanced Excel and VBA. 

### Files

* [Test Data](Resources/alphabtical_testing.xlsx) - Used this while developing scripts as it contains less data so can be tested fast.

* [Stock Data](Resources/Multiple_year_stock_data.xlsx) - For final script run.

### Stock market analyst

![stock Market](Images/stockmarket.jpg)

### Easy

* Created a script that will loop through one year of stock data for each run and return the total volume each stock had over that year.

* Displayed the ticker symbol to coincide with the total stock volume.

* Result looks as follows (note: all solution images are for 2015 data).

![easy_solution](Images/easy_solution.png)

### Moderate

* Created a script that will loop through all the stocks for one year for each run and take the following information.

  * The ticker symbol.

  * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The total stock volume of the stock.

* Used conditional formatting that will highlight positive change in green and negative change in red.

* The result looks as follows.

![moderate_solution](Images/moderate_solution.png)

### Hard

* Solution includes everything from the moderate challenge.

* Solution returns the stock with the "Greatest % increase", "Greatest % Decrease" and "Greatest total volume".

* Solution looks as follows.

![hard_solution](Images/hard_solution.png)

### CHALLENGE

* Made the appropriate adjustments to script that will allow it to run on every worksheet, i.e., every year, just by running it once.


### Other Considerations

* Used the sheet `alphabetical_testing.xlsx` while developing code. This data set is smaller and will allow to test faster. Code ran on this file in less than 3-5 minutes.

* Ensured that the script acts the same on each sheet. The joy of VBA is to take the tediousness out of repetitive task and run over and over again with a click of the button.

## Submission

* A screen shot for each year of your results on the Multi Year Stock Data.

* VBA Scripts as separate files.

