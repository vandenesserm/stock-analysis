# Stock Analysis Project

## Overview of Project
### Purpose
The purpose of this project was to refactor a VBA code to render retrieval and analysis of stock information more efficient. The initial code was developed to analyze 2017 and 2018 data from 12 stocks based on tickers. It sorted through opening and closing information, calculated volumes to determine which stocks had a positive return and which ones had a negative return. Although the code did what it was supposed to do, it did not run smoothly and took longer than what was expected.


## Analysis and Challenges
The refactoring process started with a starter code and macros created throughout the module lessons. To optimize the rendering of information, a variable called tickerIndex was created to be used across the different arrays. Using for loops to go over the spreadsheet for the yearValue in order to extract and output total daily volume and return. If-then statements were used to check for the row's ticker and assign closing prices.

The original code run time for 2017 was approximately 2.55 seconds as shown by the Msgbox. The time was calculated using the Timer function.


<img src="https://github.com/vandenesserm/stock-analysis/blob/main/PNGs/Screenshot%202022-07-08%20Code%20First%20run%202017.png?raw=true" width = 640>


The original code run time for 2018 was approximately 2.51 seconds.

<img src="https://user-images.githubusercontent.com/106083282/178146482-2cee16e3-a174-423c-9b0e-fe38ddc5eedc.png">


### Challenges and Difficulties Encountered

The refactoring process was challenging, but it was also an important learning tool. As the process spanned multiple days, the importance of a consistent workflow became evident. Much time was spent trying to understand what was previously done. By creating a consistent commenting system, I was able to quickly retrieve where I last stopped. That in itself made the whole process more efficient. 

This exercise also allowed me the opportunity to explore different avenues for assistance. Through internet research, collaboration with classmates, and tutoring assistance, having a clear understanding of the ending goal and systematic notes, allowed me to pinpoint the error sources and work through debugging.


## Results
Results

The final refactored code is as follows: 
<img src="https://github.com/vandenesserm/stock-analysis/blob/main/Code%201a-4.png?raw=true">

The refactoring process did what it was intended to do. The code was clearer, more concise, and ultimately ran much faster. The refactored code run time for 2017 was approximately 0.33 seconds.

<img src="https://github.com/vandenesserm/stock-analysis/blob/main/PNGs/Screenshot%202022-07-10%20Refactored%202017.png?raw=true">


The refactored code run time for 2018 was approximately 0.37 seconds.

<img src="https://github.com/vandenesserm/stock-analysis/blob/main/PNGs/Screenshot%202022-07-10%20Refactored%202018.png?raw=true">


## Summary
Overall the refactored code ran much faster than the original code. The code itself was easier to read and understand. Although this time the process was very laborious and difficult, with time, I believe will be an unavailable tool. Being able to constantly improve codes can have a big impact in the overall efficacy of my coding and job performance. 

The original proposal for this project was to optimize the code so it could possibly handle a larger dataset and be efficient. To that extent, refactoring is a powerful tool as it allows the betterment of current codes to handle more complex tasks and provide faster outcomes. 
