# stock-analysis
Green energy company stock analysis using VBA

## Overview of Project
The purpose of this analysis was to refactor a VBA code that was used to analyze Green Energy Stocks in years 2017 and 2018 to determine which ones were best to invest in. After refactoring the code, performance of stock analysis was assessed based on completion time and stock performance for 2017 and 2018 were also assessed.

## Results
Refactoring the original VBA script led to significantly quicker stock analysis performance run times for years 2017 & 2018 compared to the original VBA script, as intended. A comparison of analysis running times between the original and refactored scripts for 2017 and 2018 stocks are presented in the images below:

##### (Figure A) 2017 stock analysis run time from original VBA script
![Original_VBA_Challenge_2017](Original_VBA_Challenge_2017.png)
##### (Figure B) 2017 stock analysis run time from refactored VBA script
![VBA_Challenge_2017](VBA_Challenge_2017.png)

Refactored script run time = 0.09985352 seconds compared to the original script = 0.6640625 seconds. The refactored code ran the stock analysis __85%__ quicker than the original script!

##### (Figure C) 2018 stock analysis run time from original VBA script
![Original_VBA_Challenge_2018](Original_VBA_Challenge_2018.png)
##### (Figure D) 2018 stock analysis run time from refactored VBA script
![VBA_Challenge_2018](VBA_Challenge_2018.png)

Refactored script run time = 0.07910156 seconds compared to the original script = 0.6523438  seconds. The refactored code ran the stock analysis__85%__ quicker than the original code!

## Summary
### Advantages
Somes advantages to refactoring code include cleaning up the code such as removing any redundancies. This is better for reuse and to be updated by others over time. Consequently, the liklihood of encountering errors will be reduced. Refactoring also makes the code easier to read and maintain. Lastly, as demonstrated in this assignment, the analysis running time on this refactored code script was quicker because the data was only looped one time to gather the needed stock performance information.
### Disadvantages
While there are plenty of advantages to refactoring code, there are some disadvantages. Refactoring code requires extensive experience and time to avoid introducing other issues. While the make up of the code maybe improved, its functional analysis remains the same and no new functions can be added into the code as part of the refactoring prcess. 
### Advantages & Disadvanges to refactoring original VBA script
One clear advantage of refactoring the original VBA script was the improved performance in analyzing the stocks. Another advantage was looking up codes on this platform posted by previous students and I already saw how my refactored code is different where I used a nested for loop in my script while the ones I encountered just used a for loop, which appeared to look visually cleaner, and therefore easier to read and understand.  A disadvantage of refactoring is the extensive time it requires along with experience, which I am limited in both due to the nature of the boot camp.
