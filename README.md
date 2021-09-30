# VBA Challenge

## Overview of Project

This project used VBA macros in Excel to analyze green stock data.

### Purpose

The purpose of this challenge was to analyze Stock data in Excel using VBA macros. These macros were constructed such that
the client would be able to run their own analysis in the future with the click of a button, so long as the data followed the 
same format of the previous data. The refactored code was made so that the client would be able to analyze the data more efficiently. 

## Results

#### 2017

Refactored          |  Non-Refactored
:-------------------------:|:-------------------------:
![](https://user-images.githubusercontent.com/89175578/135522122-b84503ec-cb08-4ddc-b8da-c19f14cb6966.png)  |  ![](https://user-images.githubusercontent.com/89175578/135522494-9d5af31f-ea57-49af-9d98-de66046d0c9e.png)

Here are some screenshots comparing the run time of the original code and the refactored code for the 2017 year.

#### 2018

Refactored          |  Non-Refactored
:-------------------------:|:-------------------------:
![](https://user-images.githubusercontent.com/89175578/135522609-6c4fef9f-a668-4021-9c7b-3c25316a0f36.png)  |  ![](https://user-images.githubusercontent.com/89175578/135522700-0482fa43-1573-4396-a566-c228f4c982bc.png)

Here are some screenshots comparing the run time of the original code and the refactored code for the 2018 year.

### Analysis of Results

The 2017 refactored code runs in a quarter of the time as the original code, and the 2018 refactored code runs in a third of the time of the original code. Granted that these are just one test run for each of them (there is some variation with repeated testing), there was always a marked increase of speed between these two different methods of analysis, namely that the refactored always performed significantly quicker than the non-refactored code despite the fact that they produce the same results.

The major difference between these two methods is that the refactored used a ticker index that was used to run over all the code, and used arrays for the three different outputs. These outputs used the ticker index to access the stock ticker index within these arrays. 

#### Refactored Code
```
'1a) Create a ticker Index
    Dim tickerIndex as Single
    tickerIndex = 0 'set the ticker index to 0 BEFORE iterating over all the rows
    
    '1b) Create three output arrays   
    Dim tickerVolumes(11) as Long 
    Dim tickerStartingPrices(11) as Single
    Dim tickerEndingPrices(11) as Single
```
#### Non-Refactored Code
```
        For i = 0 To 11
            
            ticker = tickers(i)
            totalVolume = 0
```
The for loop used in the orginal code was also broken up into 3 for loops with an additional for loop to format the data in the same button. With the original AllStockAnalysis macro the button I made only outputs the data, but it does not format it. In the refactored code there is now a loop to initialize the starting volumes of each ticker at 0, a loop to go over all the rows in the data and store the values, a loop to output the data, and finally a loop to format the data; all in the same button.

## Summary

- Advantages and Disadvantages of refactoring code

Refactored code, when done correctly, allows one to produce the same result, but with cleaner, more efficient code. However, it was my experience that refactoring the code required a deeper understanding of the problem, the solution and the code that I struggled to grapple with. However, once these hurdles are overcome the result is a more efficient solution to the same problem, which allows the client to run their analysis more smoothly, and allows for expansion in the future. 

- Advantages and Disadvantages of refactoring code regarding this project

The obvious advantages regarding refactored code in this challenge was that it ran much faster than the original code. The disadvantage is that the refactored code would need to be altered slightly, in order to fit more tickers. However this is also an issue with the original code.
