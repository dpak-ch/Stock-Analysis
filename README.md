# Stock Analysis Using VBA

## Contents
* [Overview](#overview-of-the-project)
* [Results](#results)
* [Limitations](#limitations-of-the-refactored-code)
* [Additional Improvements](#additional-ways-to-improve-code)

## Overview of the Project
The purpose of this project is to refactor a piece of stock analysis code developed using VBA. Refactoring involves improving the code to make it more simple, readable, efficient, and consequently to remove any potential bugs. The refactored code in this project is expected to perform the same function of the original code, albeit at a faster rate.

## Results
The code calculates stock trading volume and captures opening and closing price for several stocks. The data is provided for two years (2017 and 2018) in separate spreadsheets. 

### Original Script
The original script uses a nested For statement to accomplish the task of calculating total volume over the course of a given year. The outer loop cycles through the 12 ticker symbols while the inner loop cycles through all trading data, looking for the specific ticker symbol. The image below shows the structure of the loops in vba:

![Original Script - Inner-Outer Loop](https://user-images.githubusercontent.com/81054290/115630114-18046700-a2c9-11eb-92e3-d570055e7bc4.png)

The starting price of the stock is noted at every new trigger of the inner loop. As the inner loop completes cycling through all trading data for the ticker symbol in the outer loop, the ending price for the ticker is captured. The trading volume is updated at each inner loop increment. The images below show the stock performance for two years (2017 and 2018), and the time it took for each of the runs:

![VBA_Pre-Refactor_2017](https://user-images.githubusercontent.com/81054290/115632693-e641cf00-a2cd-11eb-9def-91e50d3e0b6a.PNG)

![VBA_Pre-Refactor_2018](https://user-images.githubusercontent.com/81054290/115632701-ea6dec80-a2cd-11eb-9a02-9bc5813dd703.png)

It can be seen from the images that the green energy stocks had much better returns in 2017. While trading volume was roughly the same in both years, most of the stocks had abysmal returns in 2018. The script run times for both years is roughly the same at ~0.7 seconds.

### Modified Script for Faster Run Times
For this exercise, I followed the suggested steps to refactor the code. The original nested loop has now been replaced by two separate loops as shown below:

![Challenge Script - Individual Loops_2](https://user-images.githubusercontent.com/81054290/115775893-12fef080-a379-11eb-8dae-27433aba92e9.png)

The first loop initializes the TickerVolumes array to zero. The subsequent loop cycles through all trading data to calculate total volume for each ticker, and capture opening and closing prices. With the modified code, stock performance data remains unchanged while run time decreases by 84% from 0.7 to 0.1 seconds. The images below show stock performance and refactored code run times for 2017 and 2018:

![VBA_Challenge_2017](https://user-images.githubusercontent.com/81054290/115777917-83a70c80-a37b-11eb-9f29-75af30008238.png)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/81054290/115777807-5e1a0300-a37b-11eb-8a32-d026b4391a0c.png)

## Limitations of the Refactored code
The refactored code has accomplished the two main objectives: maintain functionality of the original code, and reduce run time. However, there is an inherent limitation not present in the original code. The refactored code assigns volume and price data to each index in the Tickers array, without checking for a ticker symbol match between the array and the spreadsheet data. The process here assumes that the Tickers array follows the order in which tickers are organized in the spreadsheet. See image below:

![Refactored Code Ticker](https://user-images.githubusercontent.com/81054290/115779486-aafed900-a37d-11eb-8c9b-1683bb6dc53e.png)

In the image above, if the user accidentally assigned tickers to the wrong array position (e.g. ticker CSIQ in position 0), then the associated volumes and prices would be wrong also. In contrast, the original code determines stock performance variables by deliberately matching ticker symbols in the spreadsheet to the Tickers array, and is hence more foolproof.

## Additional Ways to Improve Code
### Original Script
A key limitation of the original script is that the inner loop runs through all the trading data 12 times, without regard to the ticker symbol being analyzed. Some savings in time can be achieved if the inner loop is exited after a ticker match is obtained and the performance variables have been captured, as shown below:

![Original Script - Exit Inner Loop Early](https://user-images.githubusercontent.com/81054290/115782294-1eeeb080-a381-11eb-8f0d-b6ab3b6c689f.png)

![Original Script_Modified_Faster Run Time](https://user-images.githubusercontent.com/81054290/115782507-5c533e00-a381-11eb-9fd3-0bdc44f1f456.png)

While slower than the refactored code, the modified script is ~40% faster than the original script. This code is still inefficient, as the time taken to extract performance data for each subsequent ticker is higher than the previous one.

### Refactored Code
The refactored code is optimized for the task it is expected to perform. The time taken for the code to run is already low at 0.1 seconds. However, the code can be made more efficient by turning off "Screen Updating" as shown below:

![Refactored Code - ScreenUpdating Off](https://user-images.githubusercontent.com/81054290/115783580-a557c200-a382-11eb-8c87-38e9b453de9a.png)

Another small improvement was made by moving the "Returns" formatting code to the same loop where returns are calcualted as shown below:

![VBA_Challenge Script_Improved](https://user-images.githubusercontent.com/81054290/115793462-c5db4880-a391-11eb-9155-38845af6ff6a.png)

The combined improvements provide ~30% savings in run time by improving code efficieny and instructing Excel not to visibly update the spreadsheet with each change. Savings can be more dramatic with larger datasets.

![VBA_Challenge_2017_DC](https://user-images.githubusercontent.com/81054290/115784301-886fbe80-a383-11eb-88b1-1616ff9ed318.png)

![VBA_Challenge_2018_DC](https://user-images.githubusercontent.com/81054290/115785147-b4d80a80-a384-11eb-84f1-6df519a64816.png)

