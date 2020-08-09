#**Green Stock Analysis**

## **Overview of Project**

Our client, Steve, is researching stock spanning the entire market for the last few years to build a portfolio for his parents. Steve’s parents are passionate about green energy and are investing the bulk of their money in DAQO New Energy Corp (Ticker: DQ), a company that makes solar panel components. Steve would like to diversify his parents’ portfolio and has asked me to analyze DQ and several green energy stocks. 

### *Purpose of Analysis*
Steve wants to find out how actively traded (Total Daily Volume) and the percentage difference in price from the start of the year to the end of the year (Annual Return) for each stock in the portfolio in 2017 and 2018.  Using VBA, a programming language that automates tasks in Microsoft Office, our final product will allow Steve to analyze other stocks- saving time and decreasing the chance of errors. 

####**Results**

Steve’s parents believe if a stock is traded often, the price will accurately reflect the value. To test their theory, I created a for Loop to scan every row in the stock data worksheet to calculate the Total Daily Volume and Return for a selection of stocks in 2017 and 2018.  For deeper analysis, I reused code to create a flexible Macro for running multiple stocks so Steve can look at different stocks in the future. 

To make the sheet more user friendly so Steve can focus on financial analysis, I have added the following items:

- Run Analysis and Clear Worksheet buttons 

- Input Box for user to request the year they want to analyze

- Stock performance is highlighted by a loop to color cells with positive returns green and negative returns red

- Message Box to show how fast VBA code compiles results

*2017 Stock Performance*

In 2017, all but one of the selected stocks produced positive returns. DQ showed the best return of almost 200% but was traded the least (35M times). SPWR was traded the most (782M times) and had a 23% return.  The data shows Steve’s parents’ theory is incorrect this year. 

![2017 Stock Performance](https://github.com/FeliciaGanthier/Stock_Analysis/blob/master/Resources/VBA_Challenge_2017.png)

*2018 Stock Performance*

In 2018, only two of the selected stocks produced positive returns. DQ was traded 107M times and showed a -62% return. This year, Steve’s parents’ theory is still incorrect but if they followed it, they would have had a positive return. ENPH was traded the most at 607M times and showed an 82% return, 2% lower than year leader, RUN. 

![2018 Stock Performace](https://github.com/FeliciaGanthier/Stock_Analysis/blob/master/Resources/VBA_Challenge_2018.png)

*Execution Times*

The Original Script completed the 2017 analysis in 1.28 seconds and the 2018 analysis in 1.27 seconds. The Refactored Script ran in 2.34 seconds and 2.36 seconds, respectively, and the gap between the 2017 and 2018 run times increased by .02 seconds. 

#####**Summary**

*Advantages and/or Disadvantages of Refactoring Code*

Code that is clear and well-documented or readable helps developers get up to speed quickly and refactoring makes code more efficient by taking fewer steps, using less memory, or making code easier for future users.  However, each developer has their own style of using white space to organize code with spaces, tabs and line breaks and adding comments to mark sections and explain what the code is doing. For example, I like to bucket code using the Formatting, Inputs, Process and Outputs model.   For additional clarity, I add a formatting section that contains all design elements of the code such as cell coloring and font.  

*Application of Refactoring Code to the Original VBA Script*

The calculation time went up for my refactored code by approximately one second compared to the original VBA Script and the gap between the analysis execution times increased slightly. It is possible being new to VBA, I left extra code that slows down the ana
