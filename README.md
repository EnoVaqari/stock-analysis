# Stock Analysis

## Overview - Stock Analysis with VBA

### Background
Steve wanted an analysis on stock market. The analysis was performed using Virtual Basic for Application (VBA) through Excel, which created a user friendly analysis for Steve to understand the stock market.

### Purpose
Steve wanted an analysis of bot 2017 and 2018 dataset for his parents. Our goal is to edit or refactor the code, in order to get the results faster. After editing or refactoring the code, we will compare original and refactored scripts to identify the code that runs faster.

## Results

Below is the original script used to analyze the Stock Market dataset.

![](VBA_Challenge_original.vbs)

Below is the refactored VBA script used to analyze the Stock Market dataset.

![](VBA_Challenge_refactored.vbs)

### Stock Analysis for 2017 and 2018

For 2017, all stocks had a positive return on investment rate with the exeption of TERP, which had -7.21% ROI rate.
![](All_Stocks_2017.png)

For 2018, ENPH and RUN had positive outcomes with 81.92% and 83.95% return on investments respectively. The rest of the outcomes were negative. 
![](All_Stocks_2018.png)

### Original code Vs Refactored Code 

By refactoring the code, there was a visible improvement on the time the code ran from 0.9375 seconds to 0.1875 seconds for 2017, and from 0.8710938 seconds to 0.1679688 seconds for 2018. 


![](original_2017.png)
![](refactored-2017.png)
![](original_2018.png)
![](Refactored_2018.png)

## Summary

### Advantages of refactoring
Based on our analysis, refactoring  using VBA script reduced the time the code ran significantly, which is a major advantage when using large datasets.

### Disadvantages of refactoring 

Refactoring the code exposes the analysis at risk of losing or destroying parts of the original code, which can result in errors and waste of time. However, this disadvantage can be prevented by being careful, saving the original work, and comparing results to ensure accuracy.


