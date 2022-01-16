# VBA Refactored Analysis
## Overview of Project
The purpose of this analysis is to provide Steve (and his parents) a more expanded view of the data that includes the entire stock market over the last few years (2017-2018). The code in this Excel file is intended to run through a larger dataset by looping through the data over the last few years and compare the various ticker volumes and return %.
## Results 
2017 (shown below): volumes were 139,399,100 lower than they were in 2018. The % Return was largely positive across the board, with only TERP (-7.2%) showing a negative return. Overall, 2017 was a very good year for investing. 

![All Stocks_2017](https://user-images.githubusercontent.com/96449605/149681064-b9897f29-bff0-4e4a-ab4b-8605ea0fb413.png)


2018 (also shown below): trading volumes were higher, as noted above, and nearly all of the tickers that were tracked had a negative return. The only two ticker exceptions that showed a positive return were ENPH (81.9%) and RUN (84%). 

![All Stocks_2018](https://user-images.githubusercontent.com/96449605/149681104-444796f4-7612-4bb3-bff7-84291f4a7e62.png)


### Summary
The 2017 refactored code ran nearly a full second faster (0.9024 seconds) than it did compared to the regular code (shown below).
Regular code run time: ![VBA_Challenge_2017_Original](https://user-images.githubusercontent.com/96449605/149681136-372cf867-2c8b-469d-a18e-388ac5bc6ca4.png)
Refactored code run time: ![VBA_Challenge_2017_Refactored](https://user-images.githubusercontent.com/96449605/149681150-b453c80c-0720-462a-b4d9-bcc46ec70f13.png)
   
The 2018 refactored code also ran faster than the regular code (shown below) by 0.82 seconds.
Regular code run time: ![VBA_Challenge_2018_Original](https://user-images.githubusercontent.com/96449605/149681170-d03ceb43-620d-4b7d-b0cd-86730792279d.png)				
Refactored code run time: ![VBA_Challenge_2018_Refactored](https://user-images.githubusercontent.com/96449605/149681176-9dd999f4-76d8-4753-ad35-99d00dbf804a.png)
  
Advantages and disadvantages of refactoring code in general:
•	Advantages - Refactoring can refine your script to make it more efficient, both in terms of (1) the code being less clunky and also (2) making it run faster.
•	Disadvantages – Refactoring can take additional time that you don’t have and therefore may not always be an option due to time constraints. Additionally, refactoring could cause additional errors within the code and may create more problems rather than solving them.
Advantages and disadvantages of the original and refactored VBA script.
•	Advantages – The refactored code is smoother and more simplified than the original version.
•	Disadvantages – As noted above, the refactored code required additional time and resources to review, refine, debug, and run correctly.
