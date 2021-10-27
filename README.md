# refactor-analysis
Challenge assignment for Module 2

Steve loves the workbook you prepared for him. At the click of a button, he can analyze an entire dataset. Now, to do a little more research for his parents, he wants to expand the dataset to include the entire stock market over the last few years. Although your code works well for a dozen stocks, it might not work as well for thousands of stocks. And if it does, it may take a long time to execute.

In this challenge, you’ll edit, or refactor, the Module 2 solution code to loop through all the data one time in order to collect the same information that you did in this module. Then, you’ll determine whether refactoring your code successfully made the VBA script run faster. Finally, you’ll present a written analysis that explains your findings.

Refactoring is a key part of the coding process. When refactoring code, you aren’t adding new functionality; you just want to make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read. Refactoring is common on the job because first attempts at code won’t always be the best way to accomplish a task. Sometimes, refactoring someone else’s code will be your entry point to working with the existing code at a job.

Overview of Project
* The purpose of this project is to observe which green energy stock returned a positive rate or not using data from 2017 and 2018 and how fast it takes for the code to run when it has been refactored. The code would not only calculate the returns the green energy stock provided, it also measures how fast the code would work to compile the data. 

Results
* In 2017, the only stock to not return a positive result is TERP with a -7.2% drop. Meanwhile, the highest returning return is DQ with a 199.4% return. In 2018, the only stocks with positive returns were ENPH and RUN with 81.9% and 84% returns respectively; meanwhile, the biggest drop was DQ with a -62.6%.
* With a refactored code, compiling the data took 0.15625 and 0.1445313 seconds for 2017 and 2018 respectively; meanwhile, using the unrefactored code took 0.6909375 and 0.59375 seconds for 2017 and 2018 respectively.    

![2017 Runtime w/Refactored Data](https://github.com/Itgotworse26/refactor-analysis/blob/main/resources/VBA_Challenge_2017.png)

![2017 Runtime without/Refactored Data](https://github.com/Itgotworse26/refactor-analysis/blob/main/resources/VBA_Challenge_2017_No-Refactor.png)

![2018 Runtime w/Refactored Data](https://github.com/Itgotworse26/refactor-analysis/blob/main/resources/VBA_Challenge_2018.png)

![2018 Runtime without/Refactored Data](https://github.com/Itgotworse26/refactor-analysis/blob/main/resources/VBA_Challenge_2018_No-Refactor.png)


Summary
* The biggest advantage of using the unrefactored code is that it doesn't have as many variables. 

...

3c) Get the number of rows to loop over
   RowCount = Cells(Rows.Count, "A").End(xlUp).Row

   '4) Loop through tickers
   For i = 0 To 11
       ticker = tickers(i)
       totalVolume = 0
       
       '5) loop through rows in the data
       Worksheets(yearValue).Activate
       For j = 2 To RowCount
       
           '5a) Get total volume for current ticker
           If Cells(j, 1).Value = ticker Then

               totalVolume = totalVolume + Cells(j, 8).Value

           End If
           
           '5b) get starting price for current ticker
           If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               startingPrice = Cells(j, 6).Value

           End If

           '5c) get ending price for current ticker
           If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               endingPrice = Cells(j, 6).Value

           End If
       Next j

...

* The for loop for the unrefactored