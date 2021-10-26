# Module 2 Challenge
# Refactoring VBA Code for Stocks Analysis

## Overview of Project: The stated goal for this challenge is to most efficiently capture stock ticker data for our client, allowing him to recommend the best stock or basket of stocks for investment. The underpinning idea of this project is to reformat existing code evaluating these stock tickers in an effort to make improvements on the efficiency of the code. The initial stock ticker analysis macro had to run through all rows and columns of the stock data many times and the goal for our refactored code was to require only one pass through the data set, using arrays to store our eventual output values. 

 

## Results: Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.

### 
Stock Performance Comparison: 2017 vs. 2018

Overall, these twelve stocks performed much better in 2017 than in 2018. The average annual return for this group of stocks in 2017 was about 67%, while the average annual return for this group of stocks in 2018 was -8.51%. 

Here is a bar graph showing the annual return for each stock in 2017 and 2018: 

![Stock Ticker Returns: 2017 vs. 2018](https://github.com/Tozerh/stocks-analysis/blob/main/17%20vs%2018%20Comparison.PNG)

As we can see here, only one stock "RUN" posted an overall annual return in 2018 that surpassed its number for 2017 while remaining positive both years. Another ticker of note is "ENPH," performing worse in 2017 in terms of its annual return, while its returns for both years remained positive. Outside of these two tickers, the other ten all saw negative returns in 2018, indicating a tough year for this particular basket of stocks. A small bright spot might be TERP, whose returns, while negative for both 2017 and 2018, posted less of a loss in 2018, indicating a possible turnaround. 
 
### Execution Times

The refactored code was much faster than the original VBA code for both the 2017 and 2018 datasets. The difference lies in the use of arrays to store values as the code runs through each stock ticker. 

The nested for loop in the original code,
```VBA

        For i = 0 To 11
            Ticker = tickers(i)
            totalVolume = 0
            Worksheets(yearValue).Activate
          
                For k = 2 To RowCount
            
                If Cells(k, 1).Value = Ticker Then ...

```

, requires that the entire `for` loop be executed twelve times, once for each index in the array from 0 to 11. The code is very clear and the nested `for` loop is tidy, but ultimately the refactored code is much faster. 

Here are screenshots displaying the runtime of the 2017 original and refactored macros:

![2017 Original Time](https://github.com/Tozerh/stocks-analysis/blob/main/Resources/Module%202.5.3%20-%20Original%20time%20for%202017%20Analysis.PNG)

![2017 Refactored Time](https://github.com/Tozerh/stocks-analysis/blob/main/Resources/VBA_Challenge_2017.PNG)
  

And here are the screenshots of the 2018 original and refactored macros with runtime: 

![2018 Original Time](https://github.com/Tozerh/stocks-analysis/blob/main/Resources/Module%202.5.3%20-%20Original%20time%20for%202018%20Analysis.PNG)

![2018 Refactored Time](https://github.com/Tozerh/stocks-analysis/blob/main/Resources/VBA_Challenge_2018.PNG)

Refactoring the 2017 code resulted in an 86% decrease in runtime, and the 2018 refactoring resulted in an 84% decrease in runtime for the macro. What exactly accounts for such an increase in efficiency? In the case of refactoring this VBA code, the answer is in the use arrays to only require the code to run once through the source data in order to store our output values for stock ticker name, volume, starting price, and ending price. (In the refactored code, the array to store stock ticker names is the same as the original code, so I will be focusing on volume, starting price, and ending price below.)

I first had to declare these new output arrays: 

```VBA
    ReDim tickerVolumes(12) As Long
    ReDim tickerStartingPrices(12) As Single
    ReDim tickerEndingPrices(12) As Single
```
Once declared, just a single, un-nested `for` loop, not the nested `for` loop from the original code, suffices to run through the requested dataset to store output values in these new output arrays:

```VBA
 For k = 2 To RowCount
          
           tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(k, 8).Value
           
        If Cells(k, 1).Value = tickers(tickerIndex) And Cells(k - 1, 1).Value <> tickers(tickerIndex) Then

            tickerStartingPrices(tickerIndex) = Cells(k, 6).Value
        
        End If
       
        If Cells(k, 1).Value = tickers(tickerIndex) And Cells(k + 1, 1).Value <> tickers(tickerIndex) Then

            tickerEndingPrices(tickerIndex) = Cells(k, 6).Value 
            tickerIndex = tickerIndex + 1
        
        End If
    
    Next k
```

Note that the `tickerIndex` was declared as 0 just before this `for` loop. As the `for` loop runs through the indicated data sheet in Excel, each of our output arrays captures the correct value needed and, once a given ticker name is no longer detected, increases the `tickerIndex` by 1. In increasing the `tickerIndex` by 1, the `for` loop is able to run through each ticker name on one one-through of the data sheet in Excel. Since this loop only needs to run through the data sheet once, this refactored code is *much* faster than the original. 

Further, we can gain even more efficiency by removing the conditional statement that formats our code. This would require consultation with our client to ensure that changing the text color of a given cell itself would be a satisfactory replacement for changing the background fill color of a cell. If this is given an okay by the client, then we can decrease our runtime by commenting out the conditional statements that format fill color and then adding in the following for the static formatting already in the sheet: 

```VBA
  'Can insert color formatting for text here to make the program more efficient.
    Range("C4:C15").NumberFormat = "[Green]#.##%;[Red](#.##%)"

```

Here is a screenshot of the final product for the 2018 data: 

![2018 Refactored Code with NumberFormat Changes](https://github.com/Tozerh/stocks-analysis/blob/main/Resources/Module%202.5.3%20-%20Refactored%20time%20for%202018%20Analysis%20-%20With%20NumberFormat%20color%20coding.PNG)

The increase in efficiency using this method is about 20% compared to using the conditional formatting in the refactored code. 


## Summary: In a summary statement, address the following questions.
*What are the advantages or disadvantages of refactoring code?*

1) Advantages
    - Refactored code can be more efficient, especially if the code eliminates repetition, as in the case of the output arrays we used for this challenge. 
    - Refactored code should be more easily maintained, and changes to refactored code should be able to be made in less time. 
    - Refactored code should also be more compact with fewer lines, making it, if commented correctly, more readable and easier to port to other projects. 

2) Disadvantages: 
    - Refactoring code can create problems and introduce bugs that were not there in the original code. 
    - Refactoring code isn't going to change the function of the code, so if the genesis of the problem is with the larger idea of what the code should be doing or the overall structure of the code, then refactoring will not likely remedy these concerns.  


*How do these pros and cons apply to refactoring the original VBA script?*
1) Pros
    - The refactored VBA code that I created for this module was definitely more efficient, as seen in the runtime screenshots above. 
    - The refactored VBA code is also able to create the same output in fewer lines, especially if adopting the `NumberFormat` changes described above. 
    - Refactoring this code gave me a taste of what debugging code looks like after I made a mess of the original idea behind the refactoring. 

2) Cons
    - In my case, refactoring this code required a lot of testing and revising for my new `for` loops. I had originally nested my `for` loops in the refactored code, which   caused an overflow error, as my `tickerIndex` exceeded the parameters defined for my arrays. In refactoring the code, I had to take a step back and ask if the overall structure of the code made sense. Even though I had typed out code that seemed to do the same thing as my original code in fewer lines, something was clearly wrong with the bigger picture and I must have introduced a bug or two given what was happening when I tried to run the macro. 
    
