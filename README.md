# VBA of Wall Street

## Overview of the Project
The purpose of this project is to analyze a dozen of green-energy stocks to reveal their yearly transaction volume and yearly return to facilitate the decision to choose investment-worthy stocks. In this project two years’ of stock data (i.e. 2017, 2018) from 12 green-energy companies has been analyzed using Excel VBA scripts. There were 3012 rows of data in each year’s of stock data associated with those 12 companies. The analysis has been done using VBA scripts in standard logical way and also by *refactoring* codes to make the run-time faster. The comparison of run-time has been discussed in the following paragraph.

## Results
Originally the VBA scripts were written in a straight-forward logical way where the iterations were looping around the full set of data once for each of those 12 companies. After refactoring the codes by introducing a few *array* with the company tickers as *index*, the same result has been produced with lower computation / processing time. The key code is shown in the following macro snapshot.

![Refactored Code sanpshot](/Resources/refactoredCode.png) 

![2017 stocks output](/Resources/allStocks2017.png) ![2018 stocks output](/Resources/allStocks2018.png)

The original code was written in such a way that it was looking for each of those company tickers to match across the whole dataset (i.e. 3012 rows). Refactoring the code in VBA scripts allowed the computer to produce the same outcome using lower amount of mathematical processing. As a result, the run-time for the script became significantly lower as expected. The following figure shows the side-by-side comparison of the run-time for both *2017* and *2018* dataset. 

![Run time difference](/Resources/runTimeDifference.png)

## Summary
Refactoring of the code allowed the program to store individual company required data in respective arrays (i.e., *tickerVolumes(12), tickerStartingPrices(12), tickerEndingPrices(12)*) during the same iteration loop. This yields lower computation time.

The original code seems a bit easier and pretty straight-forward to think but it takes longer computational time than the refactored code. In this project the runtime was reduced roughly 4 times after refactoring the code, which could be a significant advantage in case of larger data set analysis.
