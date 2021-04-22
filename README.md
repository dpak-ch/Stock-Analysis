# Stock-Analysis

## Contents
* [Overview](#overview-of-the-project)
* [Results](#results)
* [Limitations](#limitations-of-the-refactored-code)
* [Additional Improvements](#additional-ways-to-improve-code)

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

The first loop initializes the TickerVolumes array to zero. The subsequent loop goes through every row in the spreadsheet to calculate total volume for each ticker, and capture opening and closing prices. While the stock performance data in the refactored code remained unchanged, run time decreased by 84% from 0.7 to 0.1 seconds. The images below show stock performance and refactored code run times for 2017 and 2018:

![VBA_Challenge_2017](https://user-images.githubusercontent.com/81054290/115777917-83a70c80-a37b-11eb-9f29-75af30008238.png)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/81054290/115777807-5e1a0300-a37b-11eb-8a32-d026b4391a0c.png)

## Limitations of the Refactored code
The refactored code has accomplished the two main objectives: maintain functionality of the original code, and reduce run time. However, the code has an inherent limitation not present in the original code. The refactored code captures volume and price data for each ticker symbol in the spreadsheet based on symbol change, without regard to order of data in the Tickers array. It then assigns the captured data to the TickerVolumes array for each increment of TickerIndex. The process here assumes that the Tickers array is initialized perfectly in line with the ticker data in the spreadsheet. See image below:

![Refactored Code Ticker](https://user-images.githubusercontent.com/81054290/115779486-aafed900-a37d-11eb-8c9b-1683bb6dc53e.png)

In the image above, if the user accidentally assigned tickers to the wrong array position (e.g. ticker CSIQ in position 0), then the associated volumes and prices would be wrong also. The original code however determines the stock performance variables by deliberately matching tickers in the spreadsheet to the "Tickers" array, and hence avoids this issue. 

## Additional Ways to Improve Code
### Original Script
The key limitation of the original script is that the inner loop runs through all the rows 12 times, without regard to the ticker symbol being analyzed. Some savings in time can be achieved if the inner loop is exited after a ticker match is obtained and the performance variables have been captured, as shown below:

![Original Script - Exit Inner Loop Early](https://user-images.githubusercontent.com/81054290/115782294-1eeeb080-a381-11eb-8f0d-b6ab3b6c689f.png)

![Original Script_Modified_Faster Run Time](https://user-images.githubusercontent.com/81054290/115782507-5c533e00-a381-11eb-9fd3-0bdc44f1f456.png)

While slower than the refactored code, the modified script is ~40% faster than the original script.

### Refactored Code
The refactored code is optimized for the task it is expected to perform. The time taken for the code to run is already low at 0.1 seconds. However, the code can be made more efficient by turning off "Screen Updating" as shown below:

![Refactored Code - ScreenUpdating Off](https://user-images.githubusercontent.com/81054290/115783580-a557c200-a382-11eb-8c87-38e9b453de9a.png)

This additional line provides ~30% savings in run time by instructing Excel not to visibly update the spreadsheet with each change (see images below). Savings can be more dramatic with larger datasets.

![VBA_Challenge_2017_DC](https://user-images.githubusercontent.com/81054290/115784301-886fbe80-a383-11eb-88b1-1616ff9ed318.png)

![VBA_Challenge_2018_DC](https://user-images.githubusercontent.com/81054290/115785147-b4d80a80-a384-11eb-84f1-6df519a64816.png)

