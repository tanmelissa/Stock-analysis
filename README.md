# Stock Analysis
Use VBA to help Steve analyze stock data
## Overview of Project
This project is to help Steve analyze various stocks to help his parents make a wise decision when purchasing stock. We will use Excel and VBA to analyze this stock data to help them with their decision. We will practice refactoring code to build our coding skills, and to help our macros run quicker.
## Results
### Comparing Performance of Stocks
ENPH and RUN were the only stocks to have a positive return both in 2017 and 2018. ENPG’s rate of return decreased from 2017 to 2018 while RUN’s rate of return increased from 2017 to 2018. 

TERP was the only stock to have a negative return in both years, while the remaining stocks had a positive return in 2017 but a negative return in 2018.

Looking at only these two years does not give us enough data to safely say which stocks are the safest to invest in since it is hard to see if these dips and spikes are reflective of the stocks returns over a greater amount of time, or if they are just showing one bad (or good) year for these companies. If you had to choose from this list, I would recommend ENPH and RUN. RUN may be a safer choice in that their return is increasing, unlike ENPH whose rate of return decreased between the two years. Although RUN’s rate of return was much smaller in 2017 than ENPH’s, it was higher than ENPH’s in 2018.

<img width="146" alt="Original Code 2017 Results" src="https://user-images.githubusercontent.com/102273449/167485917-d89920fb-ec95-45cd-9a05-142c324f37d6.png">
<img width="152" alt="Original Code 2018 Results" src="https://user-images.githubusercontent.com/102273449/167485927-a5f3a6c2-eb9b-49b9-ba9d-2b93a68d1955.png">

### Comparing the Execution Times of the Original Script and the Refactored Script
#### 2017
Original script run time: <img width="155" alt="Original Code 2017 Run Time" src="https://user-images.githubusercontent.com/102273449/167486081-902b351d-dab3-466a-8aa8-9dd298492e8f.png">
Refactored script run time: <img width="162" alt="Refactored Code 2017 Run Time" src="https://user-images.githubusercontent.com/102273449/167486133-0c5dde90-4cfd-4f3f-ab86-e289a767dc34.png">
#### 2018
Original script run time: <img width="161" alt="Original Code 2018 Run Time" src="https://user-images.githubusercontent.com/102273449/167486196-b4e58dea-2a3c-4881-bb7c-f278970a27f6.png">
Refactored script run time: <img width="145" alt="Refactored Code 2018 Run Time" src="https://user-images.githubusercontent.com/102273449/167486219-35bf493d-14bc-4b1c-9c3b-d35d3770512f.png">
## Summary
### What are the Advantages or Disadvantages of Refactoring Code?
### Advantages of Refactoring Code
Refactoring code can make your macros run quicker, thereby saving you time. This can be extremely advantageous if you plan on adding more and more data to your worksheets.

Refactoring code can also make it easier for other people to understand. If you plan on sharing your code with others, this is a great way to help make it easier to understand and easier for others to modify.

Refactoring code can help it use less memory when you run it. This is helpful if you want to run multiple things at once or if your data file is particularly large.

Refactoring code can also help make your code more future-proof. You can make sure to get rid of “magic numbers” and edit your code in a way that would allow your data to grow.

### Disadvantages of Refactoring Code
Refactoring code can take time to accomplish. The steps to refactor the code may not be obvious or simple, so you may spend more time cleaning up the coding than you would save by having your code run quicker.

Sometimes, when refactoring code, you may introduce new bugs or errors than was in your original code.

The refactored code may not be as simple to understand as the original code. Although you can make it easier to edit and understand overall, the logic of a refactored code may not be as obvious as the original code.

### How Does This Apply to Refactoring the Original VBA Script?
The original script has less arrays and less variables than the refactored script.

The original script only has one array: ![image](https://user-images.githubusercontent.com/102273449/167486610-4000ff38-d93e-4911-90c5-941611a32206.png)

While the refactored script has four: ![image](https://user-images.githubusercontent.com/102273449/167486658-52cce03f-96ac-42a0-82a2-b1ebe2deb60a.png)![image](https://user-images.githubusercontent.com/102273449/167486685-42665bf5-e8a6-4d9f-9787-7dba20d214c9.png)

The original script has seven variables: 
1) startTime, 
2) endTime, 
3) startingPrice, 
4) endingPrice, 
5) RowCount, 
6) ticker, 
7) totalVolume

While the refactored script has four: 
1) startTime, 
2) endTime, 
3) RowCount, 
4) tickerIndex

If you were to add more data to this excel sheet, what you would have to update in the original code would be:
- Updating the ticker array to include the new tickers added
- Updating this For loop to be able to loop through all the tickers ![image](https://user-images.githubusercontent.com/102273449/167486942-8bdf9478-d6a3-4355-a6fd-f3206f27c98e.png)

To update the refactored code, you would need to:
- Update the ticker array to include the new tickers added
- Update this For loop to initialize the tickerVolumes to zero ![image](https://user-images.githubusercontent.com/102273449/167487039-37f4868a-4c84-4804-8a4b-76db311a35f1.png)

Although you need to change the same number of items in each code, the changes needed in the refactored code are easier to spot and more intuitive, whereas updating the For loop for the tickers in the original code could easily be passed over and forgotten. Also, it would have a bigger impact on the results than if you missed updating the For loop to initialize the tickerVolumes to zero in the refactored code.

Since this code is to analyze stocks, it would be the most helpful to keep on adding sheets with data from different years. That would increase the time that it takes to run this code. This code would benefit from refactoring because of this. It will help make the code be useful for a longer amount of time, and easier to change if more tickers were added.

The original code is more intuitive when you first look at it because it seems simpler. However, adding the arrays in place of the variables is what makes the refactored code quicker easier to update in the future should more data be added. So, although we may need to add more comments into the code to make it easier for another person to understand, it is worth the extra time and effort because of the possibility of adding more data in the future.
