# All Stocks Analysis

## Overview of Project

### Purpose
The purpose of this project was to refactor previously generated VBA code with arrays in place of hard coded variables to broaden the capability of the script. 


## Analysis and Challenges



### 2017 Workbook Run Time
![2017 Analysis Run Time](https://raw.githubusercontent.com/Mrob1995/Stocks-Analysis/main/vba_Challenge_2017.png)



### 2018 Workbook Run Time
![2018 Analysis Run Time](https://raw.githubusercontent.com/Mrob1995/Stocks-Analysis/main/vba_Challenge_2018.png)

### Challenges and Difficulties Encountered
The focus for this unit was using VBA within the context of Excel to more effectively format and analyze a given worksheet, and then to refactor that code to adapt to different stock tickers and worksheet years. The element of this unit that I ended up putting the most focus and time into was eliminating "magic numbers" representing columns of specific data and replacing them with constant variables. Doing so left me with a wildly more legible set of conditionals. Once my for loops were formatted with minimal numbers, it became much easier to diagnose issues and/or plainly see what's missing. 

## Summary
Refactoring my code had the primary benefit of increased capability. The original script, on a modern computer, ran similarly fast. What changed is that the analysis is now being carried out for the full sheet. From here, it wouldn't be too difficult to substitute out the hard coded strings representing our stock tickers to instead search the sheet and auto populate with what tickers are present. It didn't make much sense given the scope of this assignment, but it wouldn't have taken all that much longer to produce a script that could consistently take in similar stock data worksheets with the same formatting and automatically populate the same analysis. As far as drawbacks to refactoring, there's no denying that the first script was simpler to write. The inclusion of arrays to work on all tickers simultaneously was a fun challenge but it did take a few hours of tinkering. In a workplace scenario, the time savings long-term of the refactored script would be tremendous as you'd be generating entire reports at the click of a button, but had there been a steady workflow of other unrelated projects this would be a somewhat time consuming step toward automation. 