# stock-analysis
Module 2 Challenge and repository for green stock analysis as part of Data Science & Visualization class

## Overview

The goal of this project was to assist our client, a financial advisor, build a tool to analyze stock data to garner insight from historical data on popular "Green Energy" stocks. This tool allows the advisor to offer meaningful guidance to their clients about how to invest directly in the "Environmental, Social, and Governance" (ESG) space. We accomplished this by using Excel and VBA to build a user interface to gather information from spreadsheets of daily stock data of 12 different companies. Our tool organizes information by stock ticker, total volume of shares traded over the course of a year, and the annual return on investment each stock earned. The tool also allows us to test the efficiency of two different methods of analyzing the data. 

This tool is meant as a guide for a financial professional. Past performance is not a guarantee of future returns. All investment products are not FDIC insured, not bank guaranteed, and may lose value. 

## Results

Our results are as follows:

![Image 1 2017 Results](https://github.com/ipbrieske/stock-analysis/blob/main/Resources/2017%20Results.png)

![Image 2 2018 results](https://github.com/ipbrieske/stock-analysis/blob/main/Resources/2018%20Results.png)

2017 was a much better year for our green energy stocks than 2018. While in 2017 all but one of our stocks achieved an annual gain, 2018 saw all but two companies lose value. While much more information is needed to make a determination as to why this shift occurred, we can determine immediately that volatility is a factor when investing in this basket of securities. Daqo New Energy Corp (DQ), for example, experienced a dizzying 199% gain in 2017, followed by a blistering loss of 63% the following year. Further analysis would be necessary in order to determine the securities with the highest overall rate of return, but our quick analysis does point to two stocks with potential as investments. The only two companies whose stock prices rose in both 2017 and 2018 were Enphase Energy (ENPH) and Sunrun Inc. (RUN). ENPH achieved a gain of 130% in 2017 followed by an 82% gain in 2018. RUN, which barely acheived a gain of 5% in 2017 went on to lead the pack at an 84% gain in 2018, potentially signalling further growth in subsequent years. 

A retrospective analysis from March 2022 shows us that all these stocks were key winners in the 2020 tech investment boom, as well as key losers in the subsequent correction in the face of hawkish Fed policy on battling inflation. Nonetheless, volatility continued to be a major factor in considering these green energy stocks as investments.

Our two methods of analyzing this data, AllStocksAnalysis() and AllStocksAnalysisRefactored(), differed significantly in their efficiency. To make this determination, we first timed the processing speed of AllStocksAnalysis(), then refactored our code to improve efficiency. In addition to recording the time necessary to process the same data, I included a table allowing me to calculate the average runtime of a series of tests. In order to get the most accurate data, I took the average runtime of 10 calculations, for each subroutine for each year of data. Our results are as follows:

![Image 3 Runtime Averages](https://github.com/ipbrieske/stock-analysis/blob/main/Resources/Final%20Results.png)

We were able to reduce our program runtime by 69.5% just by adjusting how our code stored and printed information. Nice! Here's how we did it:

## Analysis 

First, see our initial program AllStocksAnalysis()

![Image 4 AllStocksAnalysis main loops](https://github.com/ipbrieske/stock-analysis/blob/main/Resources/AllStocksAnalysis%20main%20loops.png)

Our first attempt at this program looks at each line of the spreadsheet and records data relevant to the ticker we are examining. The information is then printed by switching back to our output worksheet and then back to our data worksheet to begin the process again. Much time is wasted reading through each line of data and asking three questions - Is this the current ticker? Is this the starting price? Is this the ending price? - each time. In addition, switching back and forth between worksheets means our computer needs to continually switch between memory locations of each worksheet, wasting valuable time. Our refactored code solves both these inefficiences. 

See our refactored AllStocksAnalysisRefactored()

![Image 5 AllStocksAnalysisRefactored main loops](https://github.com/ipbrieske/stock-analysis/blob/main/Resources/AllStocksAnalysisRefactored()%20main%20loops.png)

Our new code leverages the power of arrays to store and recall lots of information quickly. While our updated subroutine uses more lines of code, it processes the data much more efficiently. We solve the issue of asking the same three questions over and over again by nesting our final two questions within a conditional for our first. If the data is not relevant to the ticker in question, do not ask if this is the starting price or ending price. We've already saved a lot of time by omitting those inquiries. AllStocksAnalysisRefactored() can be further improved; we still inquire if the data row matches the ticker in question, and if not, continue through every line in the worksheet wastefully. A future iteration of our code should eliminate this inefficiency by exiting our for loop as soon as an ending price has been determined. We also eliminate the need to flip back and forth between worksheets by using our arrays. Each index of the array is filled with the relevant information, and only when the arrays are completely populated do we flip to our output worksheet, and print the information by again cycling through our arrays. 

## Summary

There are many tradeoffs with refactoring code. In general, spending more time on a working program will increase the quality of that program - by making the code more streamlined, more readable, faster, etc. - but can take more time to improve than is actually saved by that improvement. The savings of 0.8 seconds on a spreadsheet of this size is miniscule, almost negligible, but if multiplied over a dataset orders of magnitude larger, a savings of any time can result in vast improvements in performance speed. 

Refactoring runs the risk of damaging functional code. It can be more difficult to comminicate a complex solution to a colleague. It also stretches a programmer's creativity. Revisiting old code can reveal clever, working solutions to problems encountered many months or even years later. 
