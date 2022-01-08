# Stock Analysis Overview

<!-- the purpose and background are well defined -->

## Background

Steve originally wanted an analysis of selected stocks' performance for the year 2017 and 2018 to show to his parents in order to help them pick the best stocks. Now he is asking for an analysis for all stocks that will run reliably when using several years worth of data.

## Purpose

The purpose of this challenge is to go over the original code and refactor it to run faster in case larger datasets are used down the line. We also want to see if we can combine multiple subroutines into one. This will allow us to use fewer macro buttons. The original code's purpose was to look through two data tables (2017 and 2018) to find the volume and return for each ticker and display the results on a new tab within a data table. We also added a few formatting commads in order to call out the most desired stocks at a glance.

# Results

<!-- the analysis is well described with screenshots and code -->

The results show that our refactored code works much better than the old code. The difference between both sets of code is the use of an Index variable. We used the `tickerIndex` variable to run through the tickers and it proves to loop through data more quickly. The original code looks less complicated however the refactored code runs faster.

In addition we added a message box that quantifies the results. The `MsgBox` command creates the pop-up message and displays what we ask it to. In the screen shots below you can see both message boxes and its runtime of the code.

The message box code refers to the `yearValue` variable that we created at the start of the Subroutine. This code creates an input box in which we can tell the macro which year to run the analysis on.

## Original code

```vba
If Cells(j, 1).Value = ticker Then

    totalVolume = totalVolume + Cells(j, 8).Value

End If

If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

    endingPrice = Cells(j, 6).Value

End If
```

## Refactored code

```vba
If Cells(j - 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1).Value = tickers(tickerIndex) Then

    tickerStartingPrices(tickerIndex) = Cells(j, 6).Value

End If

If Cells(j + 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1).Value = tickers(tickerIndex) Then

    tickerEndingPrices(tickerIndex) = Cells(j, 6).Value

    tickerIndex = tickerIndex + 1

End If
```

## Message Box

```vba
MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
```

## Input Box

```
yearValue = InputBox("What year would you Like To run the analysis on?")
```

Below are two message boxes that display the runtime of our original code:

<img src="https://github.com/brown-rox20/stock-analysis/raw/main/Resources/VBA_Challenge_2017_OS.png" alt="VBA_Challenge_2017_OS.png"
width="340">

<img src="https://github.com/brown-rox20/stock-analysis/raw/main/Resources/VBA_Challenge_2018_OS.png" alt="VBA_Challenge_2018_OS.png"
width="340">

Below are two message boxes that display the runtime of our refactored code:

<img src="https://github.com/brown-rox20/stock-analysis/raw/main/Resources/VBA_Challenge_2017.png" alt="VBA_Challenge_2017.png"
width="340">

<img src="https://github.com/brown-rox20/stock-analysis/raw/main/Resources/VBA_Challenge_2018.png" alt="VBA_Challenge_2018.png"
width="340">

# Summary

## Pros of Refactoring Code

Refactoring code can help when you need to use the functions in a new way. If you have to add more commands to take analysis a step further you want to look through the code and add to it in way that doesn't cause all other commands to malfuction. Combing through carefully a second time might help to find any inefficency you might have missed. If your original data file gets larger you can refactor code to ensure the command runs faster.

## Cons of Refactoring Code

There are very few reasons to not refactor code. If you don't plan to expand on your dataset or your analysis you might not need to. Overall it is a good idea to refactor code if you are going to be working with the dataset often.

## Pros of Refactoring Our Code

Our client Steve is going to use this analysis for even more stocks and even more years. It is best to find ways to make the code fuction more efficiently. Otherwise the macros would take longer to load and could malfunction altogether.

We also had a chance to consolidate commands into one subrountine. Instead of having seperate routines for the formatting and message boxes we combined all of the code into one routine. Insteading of having to push two bottons to run an analysis and a format we now only have to hit one button to have the analysis populate.

## Cons of Refactoring Our Code

The one downside to refactoring our code is that it can get a little complex. The index variable we created took time to asses and figure out where it should go. Otherwise refactoring was the best way to get Steve what he needed.
