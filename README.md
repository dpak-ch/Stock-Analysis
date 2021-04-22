# Stock-Analysis
## Overview of the Project
The purpose of this project is to refactor the code that was developed during the course of the module. Refactoring involves improving the code to make it more simple, readable, efficient, and consequently to remove any potential bugs. The refactored code in this project is expected to perform the same function at a faster rate.

## Results
### Original Script
The original script uses a nested For statement to accomplish the task of calculating total volume over the course of a given year. The outer loop cycles through the 12 ticker symbols while the innter loop cycles through all the rows checking for the specific ticker symbol. The image below shows the structure of the loops:

![Original Script - Inner-Outer Loop](https://user-images.githubusercontent.com/81054290/115630114-18046700-a2c9-11eb-92e3-d570055e7bc4.png)

The starting price of the stock is noted at every new trigger of the inner loop. As the inner loop completes cycling through all the rows for the ticker symbol in the outer loop, the ending price for the ticker is captured. The images below show the stock performance for  two years (2017 and 2018) and the time it took for each of the runs:

![VBA_Pre-Refactor_2017](https://user-images.githubusercontent.com/81054290/115632693-e641cf00-a2cd-11eb-9def-91e50d3e0b6a.PNG)

![VBA_Pre-Refactor_2018](https://user-images.githubusercontent.com/81054290/115632701-ea6dec80-a2cd-11eb-9a02-9bc5813dd703.png)

It can be seen from the images that the green energy stocks had much better returbs in 2017. While trading volume was roughly the same in both years, most of the stocks had abysmal returns in 2018. The script run times for both years is roughly the same at ~0.7 seconds.

### Modified Script - Challenge
For this exercise, I followed the suggested steps to refactor the code. The original nested loop has now been replaced by two separate loops as shown below:

![Challenge Script - Individual Loops_2](https://user-images.githubusercontent.com/81054290/115775893-12fef080-a379-11eb-8dae-27433aba92e9.png)

The first loop initializes the TickerVolumes aray to zero and the subsequent loop goes through every row in the spreadsheet to calculate total volume for each ticker and capture opening and closing prices. While the stock performance data in the refactored code remained unchanged, run time decreased by 84% from 0.7 to 0.1 seconds. The images below show stock performance and refactored code run times for 2017 and 2018.





