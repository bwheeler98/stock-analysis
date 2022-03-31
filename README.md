# Green Stock Analysis

## Overview

### Purpose
The purpose of this project was to use VBA scripting to make analyzing stock data in Excel more efficient.  VBA allows us to write a script that performs an analysis on each stock without manually entering the same formulas for each stock.  Steve initially wanted to review 12 green energy stocks for his parents and the initial VBA code worked, but could be more efficient.  The refactored code is faster and by adding a few more array variables Steve is now able to easily expand his analysis to other stocks, not just the selected green energy stocks.  

### Results
The selected green energy stocks had a much better year in 2017 than they did in 2018.  In 2017, all but one stock made money (positive return).  The stock with the biggest return was DQ at 199.4% return.  In 2018, only two stocks had a positive return.  RUN had the best return at 84%.  In regards to the actual code, the refactored code ran a lot faster than the original code.  The reason the refactored code is better and faster is because we added three new array variables as well as the tickerIndex variable.

<img src="https://github.com/bwheeler98/stock-analysis/blob/3edc4940e127288388e7c6a7c282021e99422852/refactored_code_screenshot.png" width="250" height="250">

The refactored code for 2017 & 2018 ran in 0.125 seconds.  The oringal code for 2017 ran in 0.914 seconds and 2018 ran in 0.762 seconds, which is a lot slower.  

<img src="https://github.com/bwheeler98/stock-analysis/blob/3edc4940e127288388e7c6a7c282021e99422852/resources/VBA_Challenge_2017.png" width="250" height="250"> | <img src="https://github.com/bwheeler98/stock-analysis/blob/3edc4940e127288388e7c6a7c282021e99422852/resources/VBA_Challenge_2018.png" width="250" height="250">

### Summary
A clear advantage in refactoring code based on this project is that it will run significantly faster.  Another advantage is that refactoring code gives you the ability to easily add more objects to analyze than the original code.  A disadvantage in refactoring code is that you need to be very detail oriented.  You have to make sure that when you alter one line that it won't affect other lines further down the code.  It is important to keep the code and variables consistant throughout the script.  An advantage of the original code is that it is more simple to write and understand.  However, the scope of the original code is more narrow.  It makes sense to write a more narrowly scoped or simple code for just a few objects in order to make sure the code works correctly and accurately.  Once you confirm it works, you can then refactor the code to include more objects into the analysis.
