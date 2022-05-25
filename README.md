# **STOCK ANALYSIS**
## Overview of Project
    Steve's parents decided that they wanted to invest in green energy.  They decided on DAQO (DQ) and asked Steve, due to his background in finance, to do an analysis of the stock. In 2018, DQ stock values fell 6%.  Steve hired me to perform a more indepth analysis of all stocks for 2017 and 2018 to determine how he should deversify his parent's stock portfolio.
    
### **Results**
    - _Stock Performance_
       Using refactored code, I first  initilized an array of all the green energy stocks (tickers) in the data set. Then I created the tickerIndex variable, set it to zero, and then created three output variables (tickerVolumes, tickerStartingPrices, and  tickerEndingPrices).  Once I set up my variables, I created a For loop to initialize the tickerVolumes<**Dim tickerIndex As Integer, tickerIndex = 0**> output to zero.   Then I created another For Loop to loop through all the rows of the dataset to calculate the current stocks volume <**tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value**>, the stocks starting price <**If Cells(j - 1, 1).Value <> tickers(tickerIndex) Then  tickerStartingPrices(tickerIndex) = Cells(j, 6).Value**>, and the stocks ending price <**If Cells(j + 1, 1).Value <> tickers(tickerIndex) Then  tickerEndingPrices(tickerIndex) = Cells(j, 6).Value**>. Once VBA performed that analysis, the results were displayed in a new worksheet.
       --**What was learned from this analysis?**
            1.  All of the stocks, except TERP, performed better in 2017 than in 2018. The 2017 stocks had a 8% to 199% return.
            2. The only stocks that had a good return in 2018 were ENPH and RUN. Every other stock had a loss.

         ![image](https://user-images.githubusercontent.com/104471775/170156492-0f78c2f0-9557-4108-9898-fb8b0dc0fe15.png)
         ![image](https://user-images.githubusercontent.com/104471775/170156584-2a4ca5b6-c1a9-4d12-98fa-59b29b14352a.png)

         
       --**Original code vs Refactored Code**
            The original code worked.  It did what we asked it to do but it ran much slower and was not as efficient as the refactored code.  For example, using the original code, the 2018 analysis ran in 2.9 seconds while, the refactored analysis ran in 0.3 seconds.
            ![image](https://user-images.githubusercontent.com/104471775/170156198-29a73e10-f9e7-4e72-9c54-4ec563140fa1.png)    
            ![VBA_Challenge_2018](https://user-images.githubusercontent.com/104471775/170156042-b3fd3992-9d50-4d49-b9b2-df52cc1cb6ab.png)
        
#### **Summary**
An advantage of refactoring code is that it creates more efficient code that saves on time and in the end money. A disadvantage to refacoring code is the amount of time it may take to reuse the code.  For example I for this challenge, I kept getting an overflow error.  It took several attempts befor I was able to fix my code.


