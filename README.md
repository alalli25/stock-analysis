# All Stocks Analysis
## Overview of Project
### Purpose
The purpose of this project was to use VBA in Excel to create a summary stock performance scorecard for a client, Steve, and his parents. Steve is a recent graduate in finance and is analyzing the 2017 and 2018 stock performance for 12 publicly traded energy companies. My role was to build a macro that created a summary worksheet, labeled "All Stocks Analysis" that calculates and displays the trading volume and annual return for the 12 different stocks. The analysis can be applied to both 2017 and 2018.

## Results
### Stock performance between 2017 and 2018
After running this analysis, we can confirm that almost all energy stocks tracked in 2018 posted a negative return. In 2018, Ticker ENPH & RUN both achieved a return > 80% and almost all energy stocks realized a volume increase vs prior year. My general assumption without further research is that 2018 was a poor year for all energy stocks. Steves parents were very interested in Ticker DQ. Though DQ realized a positive 199% return in 2017, I cannot recommend this stock at this time due to it's 2018 performance, where the return realized was -62.6%.

***2018 Stock Performance***

![](/Resources/Screenshots/2018_Stock.png)

***2017 Stock Performance***

![](/Resources/Screenshots/2017_Stock.png)


### Macro performance
The Macro was executed quickly for both 2017 and 2018 inputs achieving a run time < 0.12 seconds in both cases.

***2018 Macro Performance***

![](/Resources/Screenshots/2018_Macro.png)

***2017 Macro Performance***

![](/Resources/Screenshots/2017_Macro.png)



## Summary
### Advantages and Disadvantages to refactoring code
Refactoring code is a great way decrease the time it takes to address a similar question or problem. If done with efficacy, it will allow you to create, test, and run macros much more quickly than re-writing code. Unfortunately, if not done carefully, you may run into data integrity issues by introducing bugs, and creating instability.

### Pros and Cons of refactoring code for this assignment
In this assignment, we were able to directly copy and refactor a majority of the analysis output, and formatting from a previous sub script. There were additional complexities caused by creating multiple arrays to hold and store data. The final macro for this analysis used a series of arrays *(Arrays Code)* and used for loops to repeat, extract and store information *(Loop Code)*. By creating three arrays for ```tickerVolumes, tickerStartingPrice and tickerEndingPrice```, it allowed me to use ```tickerIndex``` as my index, which I believe is a cleaner way to code for the desired outcome.

***Arrays Code***

![](/Resources/Screenshots/Arrays.png)

***Loop Code***

![](/Resources/Screenshots/Loops.png)
