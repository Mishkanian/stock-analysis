# VBA Stock Analysis

## Overview of Project
Using the VBA tools inside Microsoft Excel, I constructed an algorithm to analyze over 3,000 stock data points in a user selected year to determine Total Daily Volume and Percentage Return. The purpose of this project was to refactor the operational VBA code to run faster by collecting the data more efficiently. The implementation of this improvment required changing the code to loop through the data only once, instead of multiple times.
---

## Results
After refactoring the code, I reduced the average code run time from 0.60 seconds to 0.13 seconds, a **78% increase in efficiency**. As seen in the images below, run-time is improved regardless of the selected year (2017 or 2018).

![2017:8_runtime](https://github.com/Mishkanian/stock-analysis/blob/main/Resources/VBA_Challenge_2017:8.png) 


For convenience, buttons are added directly on the All Stocks Analysis worksheet to quickly access and compare the different code. The button "Module 2 Deliverable 1" activates the refactored code. Interestinly, **using this button decreases efficiency** by approximately 0.004 seconds as opposed to going directly inside VBA. To view the comparisons more in-depth, please open "VBA_Challenge.xlsm" and **enable macros**. In Visual Basic, the orignal subroutine is found in Module1 as Sub yearValueAnalysis(). The refactored code is found inside Module2Refactor as Sub AllStocksAnalysisRefactored().

### Challenges and Debugging Code
Of the many errors that arose through refactoring the code, the most persistent issue was *Overflow (Error 6)* on Line 104. After hours of research and trials, it was found that the Overflow issue was caused due to the omission of the tickers() array in the conditional formulas on Line 71 and Line 80.

For example, on Line 71 to check whether the current row of data is the first row with the selected tickerIndex, the **non-operational** code would have read:
```
If Cells(i - 1, 1).Value <> tickerIndex And Cells(i, 1).Value = tickerIndex Then
```

However, the **corrected code** is:
```
If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
```
![Line104Error](https://github.com/Mishkanian/stock-analysis/blob/main/Resources/Line_104_error.png) 

*(The highlighted portion of code above is Line 104, the location of the Overflow error.)*

Additionally, omitting the tickers() array on Line 102 caused the stock ticker symbols to not display properly on the "All Stocks Analysis" workseet. The **incorrect line of code below** would cause the ticker symbols to display as numbers instead of letters (Example: 0, 1, 2, 3, ...).

```
 Cells(4 + i, 1).Value = tickerIndex
 ```
The **correct** line of code to display the ticker symbols as letters is:

```
 Cells(4 + i, 1).Value = tickers(tickerIndex)
```

## Summary
Refactoring code is an important process in the overall development cycle. It improves the efficiency of an operation, which is especially important when dealing with larger sets of data or more repetitve tasks. However, the decision to refactor opertational code can have unintended consequences and detriments. For example, it can break an opertion completely or interfere with other existing code. It can also be highly time consuming to identify the causes of issues and debug code.

In the specific case of this project, refactoring the code may have been excessive due to the already quick run time (0.60 seconds) and the potentially infrequent use. However, the changes made to the existing code does enable future users to upload larger sets of data without sacrificing as much run time efficiency.
