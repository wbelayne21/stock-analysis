# stock-analysis
##Overview of Project
  The purpose of this analysis is to essentially study how active a stock is as well as investigate the yearly return. The activity of a stock is depicted in the daily/annual volume which is how many shares are traded throughout the day whereas the yearly return is obtained from the difference in percentage of stock between the beginning and end of the year. 
  In this analysis, we are helping Steve choose a profitable stock for his parents. We will be looking at the total daily volume and return of ten stocks for the years 2017 and 2018. We will utilize VBA and Excel to achieve this analysis. We will refactor our original script in order to measure performance.  
###Results: 
  When analyzing the total daily volume of the 10 stocks, we learn that some had seen a significant rise in total volume while others experienced a plunge when going from 2017 to 2018. Returns on the other hand dropped substantially in 2018 compared to a year before. DQ, Steve’s parents’ choice, was at a peak of  199% return in 2017 before plunging to -62% the following year. Overall, 2017 for all most stocks was a much profitable year than 2018 where 4/5th of the ten stocks had a negative return. 
In terms of performance, both years in the original script has a similar run time of 0.66 seconds. As for the refactored code, unfortunately, I could not overcome an overflow error despite changing the data type to double and other possible types. I sought help from learning assistants online and even they could not figure out the issue. 
####Summary
  Advantages: Essentially the refactored code is more efficient as script has fewer clitters. As mentioned in the module, refactoring is a great way of introducing oneself into a new project that coworkers might already have started. Refactoring also a great way of catching code smell: redundant and bad codes. This technique facilitates easy maintenance of codes in the future. 
  Disadvantages: Code refactoring sometimes takes longer, introduces bugs and might be costly. 
The original VBA script is well composed after refactoring and should run more efficiently. Restructuring the code will make it easy to understand and follow should we need to revisit the code years down the line. It certainly takes more time to go through the script editing codes out or adding others in. 

