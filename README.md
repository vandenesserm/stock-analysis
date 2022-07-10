# Stock Analysis Project

## Overview of Project
### Purpose
The purpose of this project was to refactor a VBA code to render retrival and analysis of stock information more efficient. The initial code was developed to analyze 2017 and 2018 data from 12 different stocks, based on tickers. It sorted through opening and closing information, caculated volumes to determine which stocks had a positive return and which ones had a negative return. Altough the code did what it was supposed to do, it did not run smoothly and took longer than what was expected.


## Analysis and Challenges
The refactoring process started with a starter code with macros created throughout the module lessons. To optmize the rendering of information, a variable called tickerIndex was created to be used accross the different arrays. Using for loops to go over the spreadsheet for the yearValue in order to extract and output total daily volume and return. If-then statements were used to check for the row's ticker and assign closing prices.

The original code run time for 2017 was approximately 2.55 seconds was shown by the Msgbox. The time was calculated using the Timer function.

<img src="https://github.com/vandenesserm/stock-analysis/blob/main/PNGs/Screenshot%202022-07-08%20Code%20First%20run%202017.png?raw=true" width = 640>

The original code run time for 2018 was approximately 2.51 seconds.

<img src="https://user-images.githubusercontent.com/106083282/178146482-2cee16e3-a174-423c-9b0e-fe38ddc5eedc.png">


### Challenges and Difficulties Encountered

The refactoring process was challenging, but it was also an important learning tool. As the process spaned accross multiple days, the importance of a consistent workflow became evident. Much time was spent trying to understand what was previously done. By creating a consistent commenting system, I was able to quickly retrieve where I last stopped. That in itself made the whole process more efficient. 
This exercise also allowed me with the opportunity to explore different avenues for assistance. Through internet research, to collaborating with classmates, and tutoring assistance, having a clear understanding of the ending goal and systematic notes, allowed me to pinpoint the error sources and wok through debuging.


## Results
Results

The final refactored code is as follow: 
<img src="https://github.com/vandenesserm/stock-analysis/blob/main/Code%201a-4.png?raw=true">

The refactoring process did what it was intended to do. The code ultimately ran much faster. The refactored code run time for 2017 was approximately 0.33 seconds.

<img src="https://github.com/vandenesserm/stock-analysis/blob/main/PNGs/Screenshot%202022-07-10%20Refactored%202017.png?raw=true">


The refactored code run time for 2018 was approximately 0.37 seonds.

<img src="https://github.com/vandenesserm/stock-analysis/blob/main/PNGs/Screenshot%202022-07-10%20Refactored%202018.png?raw=true">


## Summary
There is a detailed statement on the advantages and disadvantages of refactoring code in general (3 pt).
Overall the refactored code ran much faster than the original code. The code itself is easier to read and understand. 

There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script (3 pt).