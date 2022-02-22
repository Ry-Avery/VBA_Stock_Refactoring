# **VBA Stock Refactoring**

---

## **Overview of Project**
- Optimize a VBA macro to analyize larger stock data sets.

### Purpose
- Learning how to refractor, or edit, code is an important part of being a developer. Using a few Excel stock data sets, I wrote a subscript that pulled the total volume and yearly returns onto a separate sheet. The code I wrote was rather inefficient, so I was given the task of improving the runtime while keeping all the features my previous code had.

## **Results and Analysis**
- Overall, I succeeded in refactoring my old code, improving runtime by over a second. I ended up borrowing some of my old code and rewriting most of it. I also ended up declaring a lot more variables than my original had so that my macro would be compatible with older versions of VBA. But, by far, the most substantial improvement I made was decreasing the number of times I looped through the data. I used multiple arrays and a "tickerIndex" variable to limit the number of times I had to loop through the data sets.

### Refactored Loop

```

For r = 2 To RowCount
    If Cells(r, 1).Value = tickers(tickerIndex) Then
    'Looping through all rows in sheet and checking to see if sheet ticker matches tickerIndex
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(r, 8).Value
        'Compiling the total volume for each ticker
        If Cells(r - 1, 1) <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(r, 6).Value
        End If
        'Getting the starting price for the year for the specific ticker
        If Cells(r + 1, 1) <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(r, 6).Value
            tickerIndex = tickerIndex + 1
        End If
        'Getting the ending price and switching tickerIndex if next ticker value does not match the current one
    End If
Next r

```

### Runtime Comparaison
- It is worth noting that while I had two different data sets, they had the same number of rows to loop through, so their runtime was about the same. Previously the runtime was just over 1.2 seconds, but now my code runs around the 100-millisecond mark. This makes my refactored code twelve times faster than before, which lines up with the code I changed as my original script looped over the data twelve times more than my current script.  
![2017 runtime](/Resources/VBA_Challenge_2017.PNG)
![2018 runtime](/Resources/VBA_Challenge_2018.PNG)

## **Summary**
- **What are the advantages or disadvantages of refactoring code?**
*There are significantly more advantages than disadvantages when refactoring code. Advantages include the aforementioned improved runtime, cleaning up code to make it easier to read, and cutting out unnecessary lines of code. The only real disadvantage is the increased time spent on refactoring.*

- **How do these pros and cons apply to refactoring the original VBA script?**
*I was able to cut my runtime a lot by refactoring my code, but I ended up adding a few more lines of code declaring variables and adding explanations of what the code was doing.*
