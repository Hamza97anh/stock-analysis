# stock-analysis.

# Stock Analysis Via Excel VBA

## Purpose
In this project we refactor the stock market dataset with VBA. Using VBA we loop throught the data and compile the entire dataset into a digestable table that showcases the total daily volume in comparision to it's percent return. Once a benchmark is set, we improve out code by refactoring it and making it faster. By the end of this, you will see that we were able to acheive the same results more efficiently. 

## Background

 "Steve loves the workbook you prepared for him. At the click of a button, he can analyze an entire dataset. Now, to do a little more research for his parents, he wants to expand the dataset to include the entire stock market over the last few years. Although your code works well for a dozen stocks, it might not work as well for thousands of stocks. And if it does, it may take a long time to execute.

In this challenge, you’ll edit, or refactor, the Module 2 solution code to loop through all the data one time in order to collect the same information that you did in this module. Then, you’ll determine whether refactoring your code successfully made the VBA script run faster. Finally, you’ll present a written analysis that explains your findings.

Refactoring is a key part of the coding process. When refactoring code, you aren’t adding new functionality; you just want to make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read. Refactoring is common on the job because first attempts at code won’t always be the best way to accomplish a task. Sometimes, refactoring someone else’s code will be your entry point to working with the existing code at a job."

## Results: Refactored VBA Code

1- Created a ticker Index, and three output arrays

![VBA_Challenge resource 1](https://github.com/Hamza97anh/stock-analysis./blob/main/VBA_Challenge%20resource%201.png)

Created this tickerIndex variable and set it equal to zero. This allows the code to run through all the rows. Then we created three output arrays which is where the data will land.

2- Created a for loop to intialize the tickerVolumes to zero.

![VBA_Challenge resource 2](https://github.com/Hamza97anh/stock-analysis./blob/main/VBA_Challenge%20resource%202.png)


Activated worksheet with yearValue variable. TickervVolumes is set to zero. Then we loop over all the rows in the user specified worksheet.

3- Created Script that loops through the stock data.

![VBA_Challenge resource 3](https://github.com/Hamza97anh/stock-analysis./blob/main/VBA_Challenge%20resource%203.png) 

The script collects data for tickers, tickervolumes, tickersStartingprice, and tickerEndingPrices. The script increases the current tickerVolumes variable every time it spots the same ticker. Then to give value to the tickerSartingPrices and tickerEndingPrices we created a If Then formula that checks to see if the value before the cell is greater then or less then the tickerindex and the current cell is equal to the tickerIndex then it will store it as the tickerStrartingPrices. To find the tickerEndingPrice we check the cell after if it's greater then or less then to which that value becomes the tickerEndingPrices.

4- Created the script for outputting the value into a table with assigned headers.

![VBA_Challenge resource 4](https://github.com/Hamza97anh/stock-analysis./blob/main/VBA_Challenge%20resource%204.png)

The table is color coded and bolds the column headers for a tidy look. The numbers are displayed in standard form thanks to formating script that automaticlly makes the cells fit the values. 

## The output tables 

 The refactored and the orginal for comparion are very similar and display rough the same values. 
 
 Refactored 2017
 
 ![VBA_Challenge 2017 refactored](https://github.com/Hamza97anh/stock-analysis./blob/main/VBA_Challenge%202017%20refactored.PNG)
 
 Original 2017
 
 ![VBA_Challenge 2017](https://github.com/Hamza97anh/stock-analysis./blob/main/VBA_Challenge%202017.PNG)
 
 Refactored 2018
 
 ![VBA_Challenge 2018 refactored](https://github.com/Hamza97anh/stock-analysis./blob/main/VBA_Challenge%202018%20refactored.PNG)
 
 Original 2018
 
 ![VBA_Challenge 2018](https://github.com/Hamza97anh/stock-analysis./blob/main/VBA_Challenge%202018.PNG)

### Time Lapse Results

For 2017

![VBA_Challenge_2017](https://github.com/Hamza97anh/stock-analysis./blob/main/VBA_Challenge_2017.png)

For 2018

![VBA_Challenge 2018](https://github.com/Hamza97anh/stock-analysis./blob/main/VBA_Challenge_2018.png)

# Summary

## 1. What are the advantages or disadvantages of refactoring code?

### Advantages:
- Improving the end product by either finding glitches, or making it run smoother and more efficent.
- Helps to go back and reorganize and write down the logic steps for each script for ease of editing in the future. Sometimes we'll be to overwhelmed in the begning when faced with the task of writing a code from scratch that we may forget to label or writting a unneccasry code. 
- You can discover new functions of your programing language that you can use to be more efficent in the future.

### Disadvantages:
- Refactoring a code can create altered results. In our case the numbers were accurate but not exact. If you're doing account you want the numbers to be perfect
- It can be time consuming going over you're old script if the sript wasn't labeled properly and you have reverse engineer your code. 

## 2. How do these pros and cons apply to refactoring the original VBA script?

Consider any code you write as a draft. There is always room for improvment specially when writing a long code. Refactoring helps clean up and organize your script. The nature of technology and coding is the goal of automating processes and saving time. Being able to go back and find flaws easily with your code is always going to be a usuaful tool in the industry. 
