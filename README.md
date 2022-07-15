#  Stock Analysis Refactoring

## Overview of Project
A visualization and interpretation of refactoring VBA code to run faster.

### Purpose
Using a curated stock price dataset that includes fields such as dates, volumes, and prices, this report will attempt to show how refactoring VBA code makes it more efficient.  The focus of the report will be refactoring code running against the years 2017 and 2018.

---

## Results

### Stock performance from 2018 compared to 2017
1. Only two stocks increased in value from the end of 2017 to the end of 2018.  One possible reason that many of the energy stocks in this selection lost value in 2018 was due to energy policies of the Trump Administration being hostile towards alternative forms of energy production.

![Pre-Refactor 2017](/green_stocks_2017.PNG)

![Pre-Refactor 2018](/green_stocks_2018.PNG)

2. The amount of time it took to complete the script after refactoring was significantly less; 10 times as fast, as a matter of fact.  This is due to the reduction in unneeded iterations.

![Post-Refactor 2017](/VBA_Challenge_2017.png)

![Post-Refactor 2018](/VBA_Challenge_2018.png)

The real secret here is the removal of the unnecessary For loop that forces the code to loop through the dataset 12 times.  The refactored code only loops through the dataset once, but it still collects all of the pertinent data that we need to display.

```
    '4) Loop through tickers
    For i = 0 To 11
        ticker = tickers(i)
```

versus

```
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(i) = 0
    Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
```

---

## Summary

### What are the advantages or disadvantages of refactoring code?
1. One advantage of refactoring the code is that it reduces the number of computational cycles considerably.  
2. Another advantage of refactoring the code is that it reduces the amount of actual text that a programmer needs to read through.

### How do these pros and cons apply to refactoring the original VBA script?
1. For smaller datasets, this may not be too important, but for larger datasets, shaving minutes (or seconds!) off of computations allows stock traders to make decisions in quasi-realtime.
2. Refactoring the code makes it more efficient, and allows subsequent programmers who touch the code to understand what the purpose of the procedure was much faster.
