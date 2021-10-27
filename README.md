# refactor-analysis
Challenge assignment for Module 2

In the future, Steve may want to perform his analysis on larger datasets, and he wants to know how fast his VBA code will compile the results. To help Steve, we need to add a script that will calculate how long the code takes to execute and output the elapsed time in a message box.

Overview of Project
* The purpose of this project is to observe which green energy stock returned a positive rate or not using data from 2017 and 2018 and how fast it takes for the code to run. The code would not only calculate the returns the green energy stock provided, it also measures how fast the code would work to compile the data. 

Results
* In 2017, the only stock to not return a positive result is TERP with a -7.2% drop. Meanwhile, the highest returning return is DQ with a 199.4% return. In 2018, the only stocks with positive returns were ENPH and RUN with 81.9% and 84% returns respectively; meanwhile, the biggest drop was DQ with a -62.6%.
* With a refactored code, compiling the data took 0.1445313 and 0.1523438 seconds for 2017 and 2018 respectively; meanwhile, using the unrefactored code took 0.59375 and 0.6909375 seconds for 2017 and 2018 respectively.    

![VBA_Challenge_2017.png]
![VBA_Challenge_2018.png]


Summary
* 