# Visual Basic Stock Analysis

## Overview

The purpose of the VBA Stock Analysis challenge was to refactor the VBA code that I worked on in the module lessons to make it more efficient. The original VBA code contained nested loops, which worked well for a small amount of data, but could run into efficiency issues with a larger dataset. In this challenge, I was tasked with creating a summary of 12 stocks containing ~3k lines of data that would run through VBA and provide a summary of each of the stocks perfomance for 2017 and 2018 depending on what the user inputs.

## Results of Refactoring the VBA Code

Although the refactoring process was challenging, I was very impressed with the results. The execution times from the refactored scripts compared to the original execution times processed about 6x faster - I think this performance gap would actually increase beyond 6:1 times faster if the dataset was even larger.

Original Run Times:

https://github.com/jasonatoledo/stocks-analysis/blob/master/Resources/Original_2017_Macro.png and https://github.com/jasonatoledo/stocks-analysis/blob/master/Resources/Original_2018_Macro.png

Versus Refactored Run Times:

https://github.com/jasonatoledo/stocks-analysis/blob/master/Resources/VBA_Challenge_2017.png and https://github.com/jasonatoledo/stocks-analysis/blob/master/Resources/VBA_Challenge_2018.png

One of the challenges with this challenge was assigning arrays - This wasn't explicitly explained in the actual modules, so I had to do some research which led to the code below:

Dim tickerVolumes(12) As Long

Dim tickerStartingPrices(12) As Single

Dim tickerEndingPrices(12) As Single


## Summary:

### Advantages/Disadvantages to refactoring code

The primary advantages to refactoring code is that the code can run a lot faster and many times will be more readable by other users. In addition, refactoring the code can help find bugs or inefficiencies in the code.

The disadvantages include time consumption and the frustration that comes with debugging updates. Sometimes if one block of code is improved, another may break. Another factor to that can be a disadvantage is if the refactored code only provides a minimal boost in efficiency or if the product/code won't be used often enough to gain any benefit from the time spent refactoring the code. Worst case scenario, the analyst/engineer refactoring the code could possibly break something that wasn't intended and isn't completely obvious until a user experiences the bug.

### How the above pros and cons applied to refactoring this project

The pros apply to refactoring the original VBA script in that the results came out 6x faster and doesn't have nested loops where the code has to iterate through almost 10x as many rows in order to process the results.

The con for refactoring the original code is that I spent close to 8 hours trying to figure out how to do this (I know it is a learning experience!) and I'm certain with more practice I would become much faster at writing the code. However, as mentioned as a con, this refactoring effort will only be used on 1 dataset, so if this was a real-world scenario, it would be considered a waste of resources if the analyst were to work on a tool/macro that would only be used one time.
