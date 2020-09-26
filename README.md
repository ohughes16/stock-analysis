# stock-analysis

## Overview of Project
I wrote Steve code to analyze various workshets of stock data based on year at the click of a button. The stock ran well with the 12 tickers for 2017 and 2018, with about 3012 rows of data to loop through. The code consisted of a nested for loop. This required the macro to run through each of the 3012 rows 12 times to collect the data that Steve needed. 

I wanted to investigate if the code could be designed to run faster if we were to increase the number of tickers being analyzed down the line. This project was designed to compare the efficiency of different VBA logic for Steve's analysis. If we took the nested for loop out of the macro, and were able to write the macro to have a single for loop with values stored in an array and used a counter to index the data as we went through the sheet one time. 
## Results
Utilizing a single for loop to go through the 3012 rows a single time caused the speed of the macro to increase by a factor of 10.



## Summary
### What are the advantages or disadvantages of refactoring code?
### How do these pros and cons apply to refactoring the original VBA script?
