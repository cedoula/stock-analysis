# **Stocks Data Analysis**

## Overview of Project

### Purpose
The purpose of this analysis is to provide our client Steve with an Excel workbook including an easy-to-run VBA macro able to analyze an entire dataset of stocks.\
The analysis was ran using two VBA scripts: the original script produced through the Module #2 and a refactored version of it. We would then have the opportunity to compare the performances of both scripts, and highlights the pros and cons of refactoring a code.

## Results

### Stock Performance Comparison Between 2017 and 2018

These are the screenshots of the stocks performance in 2017 and 2018 obtained with both scripts:

----Picture stocks results for 2017 | Picture stocks results for 2018

In 2017, all stocks except TERP have a positive return. 4 stocks even have over 100% return: DQ, ENPH, FSLR and SEDG. The best performing one was DQ with 199.4% return!\
This was a performing year for the green stocks.\
\
In 2018 mostly all green stock returns went in the red. Only ENPH and RUN continued to perform with respectively 81.9 and 84.0% that year.\
ENPH is globally the best performing green stock over 2017 and 2018.\

### Performance Comparison Between Original and Refactored Scripts
The analysis was ran with the original VBA script obtained through Module #2 and the refactored script for the Challenge #2.
Let’s compare the execution times of both scripts.\

---- Pictures timer for original | Pictures timer for refactored  
****put image titles
—— Pictures of code??? original and refactored..


Refactoring the script increase tremendously its performance as seen on the screenshots above the execution time gets shorter.

## Summary

1. What are the advantages or disadvantages of refactoring code?
	- pros: make code faster, less variables declared so less memory used, clearer code, easier to read and modify for future updates, more flexible code
	- cons: no additional functionality, basically doing same thing faster, can bring more complexity

2. How do these pros and cons apply to refactoring and the original VBA script?
	- execution time is faster but we are still analyzing the 12 set of green stocks. Refactored script does not give ability to analyze any or the whole set of existing stocks. 
