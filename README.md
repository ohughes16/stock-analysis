# stock-analysis

## Overview of Project
I wrote Steve code to analyze various years of stock data at the click of a button. The original VBA script ran well with the 12 tickers for 2017 and 2018, with about 3012 rows of data to loop through. The original code consisted of a nested for loop. This required the macro to run through each of the 3012 rows 12 times to collect the data that Steve needed for his analysis.

I wanted to investigate if the code could be designed to run faster for the purpose of potentially increasing the number of tickers being analyzed in the future. This project was designed to compare the efficiency of different VBA logic for Steve's analysis. If we took the nested for loop out of the macro and were able to write the macro to have a single for loop this would mean only having to run through the spreadsheet of 3012 rows a single time. 
## Results
Utilizing a single for loop to go through the 3012 rows a single time resulted in the speed of the macro to run 12 times faster. We utilized a timer at the begining and the end of the code to compare the speed of the original and refactored macros. Below is a visual of the time that it took the original and refactored code to run for each year that I analyzed.

![VBA_Challenge_2017_original](https://user-images.githubusercontent.com/64506842/94330506-229c8f00-ff7a-11ea-84d4-b6b059371a56.PNG)
![VBA_Challenge_2017_refactored](https://user-images.githubusercontent.com/64506842/94330509-23352580-ff7a-11ea-92bd-0a5177e5a04a.PNG)

![VBA_Challenge_2018_original](https://user-images.githubusercontent.com/64506842/94330510-23cdbc00-ff7a-11ea-8e42-d2377cb2ce54.PNG)
![VBA_Challenge_2018_refactored](https://user-images.githubusercontent.com/64506842/94330511-24665280-ff7a-11ea-8fa4-097bed40a160.PNG)

## Summary
### What are the advantages or disadvantages of refactoring code?
The refactored code runs through the spreadsheet of the 12 tickers 12 times faster than the original VBA script. As the number of tickers that need to be analyzed increase, the speed of the code would remain much faster than the original VBA script. The disadvantage to the refactoring of this code is that the tickers must be in order for the macro to run correctly.  If the tickers were not incremental, the values would be indexed in the array incorrectly. 
### How do these pros and cons apply to refactoring the original VBA script?
The benefit to the original VBA script is that the tickers do not have to be sequential for the outcome to remain to be correct. Since you're running through the entire sheet looking for each ticker individually, you are able to look through ticker data that is out of order. 
The disadvantage of the nested for loop is that as the number of tickers increases, the for loop has to run through the entire spreadsheet that many  more times.
