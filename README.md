# VBA-challenge
VBA Homework - The VBA of Wall Street

For this challenge, I used two scripts:
  Module2.bas (stock_summarizer) reproduces the desired results as per the instructions.
    1. Loops through each spreadsheet.
    2. Creates a summary table with four columns - Ticker, Yearly Change, Percent Change, and Total Stock Volume.
      (a) Lists unique ticker symbols in column J.
      (b) Computes the difference between opening (on day 1) and closing (years-end) values of the stock. Negative change is color-coded RED and positive/zero change GREEN. 
      (c) Computes the difference in (2) above as a percent change as compared to the opening value and formats the value as a percentage.
      (d) Computes the total stock volume (sum of column G for each stock).
    
  Module3.bas (Clear_Summary) resets the spreadsheet to remove the summary table contents and formatting.
  
The images are before and after screenshots of the workbook.

The workbook has two buttons to call the scripts. One generates the summary (GENERATE SUMMARY) and the othe resets the workbook (CLEAR SUMMARY). 
The workbook is accessible here https://drive.google.com/file/d/1rnPWNZc_vTrQV1R4uWkfzuGyc4TxF7rJ/view?usp=sharing
