# Overview of VBA Stock Analysis Project
## Editing and Refactoring Excel VBA Code
In this challenge we delved into the code we’ve been working on in the module, except changes were made with missing aspects for us to figure out, with the ultimate end goal of making the code more readable and efficient, all whilst delivering the exact same results as before.
## Results and Comparison To Previous Code
##### Stock Performance Analysis
In terms of the end results for both versions of the code, they delivered the exact same end result. What we see is the market performance of 12 different stocks for 2 separate 1 year periods; 2017 and 2018. For both years we have the sum of their total daily trading volume for the year, and the overall return for that year. It’s clear that 2017 was a much better year, as the majority of stocks in the sample flourished and yielded exceptionally high returns for a single period, whereas 2018 was the opposite, however to a much less extent.
##### VBA Script Run Time Comparison
Our refactored/edited code ran in ≈ 0.09 seconds for 2017, and ≈ 0.10 seconds for 2018. When returning to our original code 2017 ran at ≈ 0.91 seconds, and 2018 ≈ 0.83 seconds. While this is insignificant in the big picture, it’s worth noting specific changes in the newer code that ultimately caused this. One thing worth noting is the usage of variables. The original code had used both (i) and (j) for the loops and conditionals, where the new code only used (i). The new code also utilized more efficiency with the conditionals referring to start and end price of each stock. Before, it had to be declared if the cell variable was not equal to the cell 1 row before/after it, and also if the cell displayed the ticker it was running the loop through. Our new script removed the “and equal to” aspect to the loops for starting/ending price, while ending with the same result, on a slightly more efficient, simpler path.
**Example of before:**
```
        If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
               startingPrice = Cells(j, 6).Value
        End If
```
**Example of after:**
```
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
               tickerStartingPrices(tickerIndex) = Cells(i, 6).Value       
        End If
```
## Pros and Cons of Refactoring the VBA Code
##### Cons
The problem with the editing of this code is going back and changing things leads to a bigger margin for error, when realistically there wasn’t anything wrong with it in the first place. I encountered a single minor spelling error when refactoring, which was rather time consuming to eventually spot it when VBA couldn’t, which didn’t allow it to run properly. As well as realizing an (i) variable had to be added to an array declarations that didn’t need to before, this added an element of frustration, especially knowing everything worked fine before all these changes.
##### Pros
The obvious upside to updating already functional code is it’s efficiency/run time, which is again insignificant in this example, but for much larger datasets and projects it can make a big difference, especially for it’s ability to be understood by other people, and the things learned in the process to make code more efficient in the future.

