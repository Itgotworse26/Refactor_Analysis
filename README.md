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
* The biggest advantage of using the unrefactored code is that the for loop hard codes the cells. As seen in the code below, the cells are hard coded according to whether their index matches what the code is looking for, whether it is the total volume, starting price, or ending price. This makes it easier to track and assess what populates each cell, which makes it easier for non-coders to follow.

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

* The for loop for the refactored code on the other hand has to check the tickerIndex has the correct row. While tickerIndex is the only variable that needs to be checked and compared against the cells, it can be difficult for non-coders to track the logic, especially as how tickerIndex's value can affect the tickerVolumes, tickerStartingPrices, and tickerEndingPrices depends on how many loops it has gone through; trying to track the number of loops however is difficult without machine assistance. 

...

    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
    
    Next i
    
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        
        End If
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next rowâ€™s ticker doesnâ€™t match, increase the tickerIndex.
        'If  Then
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        
        End If

            '3d Increase the tickerIndex.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
            
            End If
            'End If
    
    Next i
...

* However, refactored code will run through any calculations faster than unrefactored code. In the refactored code, tickerIndex is the only variable that the for loop increases on its own, with tickerVolumes, tickerStartingPrices, and tickerEndingPrices being affected by tikerIndex. Meanwhile, the unrefactored code needs a nested for loop and needs to check every single cell. This makes it slower and eats up more memory. While the differences are negligible at this stage due to the small dataset, a larger dataset would see a dramatic gap in performances between the two different codes. 