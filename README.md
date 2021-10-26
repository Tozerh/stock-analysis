# Module 2 Challenge
# Refactoring VBA Code for Stocks Analysis

## Overview of Project: The stated goal for this challenge is to most efficiently capture stock ticker data for our client, allowing him to recommend the best stock or basket of stocks for investment. The underpinning idea of this project is to reformat existing code evaluating these stock tickers in an effort to make improvements on the efficiency of the code. The initial stock ticker analysis macro had to run through all rows and columns of the stock data many times and the goal for our refactored code was to require only one pass through the data set, using arrays to store our eventual output values. 

 

## Results: Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.

### 
Stock Performance Comparison: 2017 vs. 2018

Overall, these twelve stocks performed much better in 2017 than in 2018. The average annual return for this group of stocks in 2017 was about 67%, while the average annual return for this group of stocks in 2018 was -8.51%. 

Here is a bar graph showing the annual return for each stock in 2017 and 2018: 

![Stock Ticker Returns: 2017 vs. 2018](https://github.com/Tozerh/stocks-analysis/blob/main/17%20vs%2018%20Comparison.PNG)

As we can see here, only one stock "RUN" posted an overall annual return in 2018 that surpassed its number for 2017 while remaining positive both years. Another ticker of note is "FLSR," performing worse in 2017 in terms of its annual return, while its returns for both years remained positive. Outside of these two tickers, the other ten all saw negative returns in 2018, indicating a tough year for this particular basket of stocks. A small bright spot might be TERP, whose returns, while negative for both 2017 and 2018, posted less of a loss in 2018, indicating a possible turnaround. 
 
### Execution Times

The refactored code was much faster than the original VBA code for both the 2017 and 2018 datasets. The difference lies in the use of arrays to store values as the code runs through each stock ticker. 

The nested for loop in the original code,


        For i = 0 To 11
            Ticker = tickers(i)
            totalVolume = 0
            Worksheets(yearValue).Activate
          
                For k = 2 To RowCount
            
                If Cells(k, 1).Value = Ticker Then ...
                

, requires that the entire for loop be executed twelve times, once for each index in the array from 0 to 11. The code is very clear and the nested loop is tidy, but ultimately the refactored code is much faster. 

Here are screenshots comparing the runtime of the 2017 original and refactored code:

![2017 Original Time](https://github.com/Tozerh/stocks-analysis/blob/main/Resources/Module%202.5.3%20-%20Original%20time%20for%202017%20Analysis.PNG)

![2017 Refactored Time](https://github.com/Tozerh/stocks-analysis/blob/main/Resources/VBA_Challenge_2017.PNG)
  

And here are the 2018 original and refactored screenshots with runtime: 

![2018 Original Time](https://github.com/Tozerh/stocks-analysis/blob/main/Resources/Module%202.5.3%20-%20Original%20time%20for%202018%20Analysis.PNG)

![2018 Refactored Time](https://github.com/Tozerh/stocks-analysis/blob/main/Resources/VBA_Challenge_2018.PNG)

Refactoring the 2017 code resulted in an 86% decrease in runtime, and the 2018 refactoring resulted in an 84% decrease in runtime for the macro. So...what exactly accounts for such an increase in efficiency? In the case of refactoring this VBA code, the answer is in the use arrays to only require the code to run once through the source data in order to store our output values for stock ticker name, volume, starting price, and ending price. In the refactored code, the array to store stock ticker names is the same as the original code, so I will be focusing on volume, starting price, and ending price below. 

I first had to declare these new output arrays: 

'''
    ReDim tickerVolumes(12) As Long
    ReDim tickerStartingPrices(12) As Single
    ReDim tickerEndingPrices(12) As Single
'''


## Summary: In a summary statement, address the following questions.
What are the advantages or disadvantages of refactoring code?

Group logically related data together – let’s say you want to store a list of students. You can use a single array variable that has separate locations for student categories i.e. kinder garden, primary, secondary, high school, etc.
Arrays make it easy to write maintainable code. For the same logically related data, it allows you to define a single variable, instead of defining more than one variable.
Better performance – once an array has been defined, it is faster to retrieve, sort, and modify data.
How do these pros and cons apply to refactoring the original VBA script?
