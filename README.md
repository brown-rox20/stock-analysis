# Stock Analysis Overview

<!-- the purpose and background are well defined -->

## Background

Steve originally wanted an analysis of selected stock's performance for the year 2017 and 2018 to show to his parents in order to help them pick the best stocks. Now he is asking for an analysis for all stocks that will run reliably when using several years worth of data.

## Purpose

The purpose of this challenge is to go over the original code and refactor it to run faster in case larger datasets are used down the line.

# Results

<!-- the analysis is well described with screenshots and code -->

The original code's purpose was to look through two data tables(2017 and 2018) to find the volume and return for each ticker and display the results on a new tab within a data table. We also added a few formatting commads in order to call out the most disired stocks at a glance.

The difference between the orginal code and the refactored code is the Index variable used to run through the tickers. The tickerIndex seems to loop through data quicker. You can see below both iterations of the code, index vs. no index.

## Original code

```
If Cells(j, 1).Value = ticker Then

    totalVolume = totalVolume + Cells(j, 8).Value

 End If

If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

    endingPrice = Cells(j, 6).Value

End If
```

## Refactored code

```

If Cells(j - 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1).Value = tickers(tickerIndex) Then

    tickerStartingPrices(tickerIndex) = Cells(j, 6).Value

End If

If Cells(j + 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1).Value = tickers(tickerIndex) Then

    tickerEndingPrices(tickerIndex) = Cells(j, 6).Value

    tickerIndex = tickerIndex + 1

End If
```

The original code looks less complicated however the refactored code runs faster. In the screen shots below you can see a message box displays the runtime of the original code. The MsgBox command creates the pop up in relation to the yearValue

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

<!-- there is a detailed statement on the advantages and disadvantages of refactoring code in general
there is a detailed statemnent on the advantages and disadvantages of the original and refactored VBA script --!>
