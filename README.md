# Stock_Analysis

## Overview of Project

### Purpose

We built a workbook for steve already but now we have to do a little more research for his parents. Our job is to expand the dataset to include entire stock market over the last few years. We have built a faster code that can run the entire stock dataset in about 0.125 seconds which helps us and is convenient for both us and the readers.

## Results

### Findings
 So from what we've understood in 2017 had a way better performance than 2018 with using (AY) as an example AY had a Total daily volume of 136,070,900 with a return of 8.9% while in 2018 AY dropped to a whopping total daily volume of 1,247,400 and a 0% return which is a huge drop in performance than 2017. I have included a photo of the 2017 picture here 
 [![VBA_Challenge_2017](Resources/VBA_Challenge_2017.png)
 
 And, You are also able to compare with the 2018 image here to see the differences
 [![VBA_Challenge_2018](/Resources/VBA_Challenge_2018.png)

Also there is also a difference in exucution times between the two years, in 2017 it ran in 0.121 wheras in 2018 it runs in 0.085

### Coding 

I Created an index for the tickers and created three output arrays after that i had to make the volumes of the tickers = 0 to 11 even though there are 12 tickers as our string. This is because the code always starts at 0 and 0 is considered a value here after we finished that we created loops for the indexes and we check for the correct rows and tickers. We ended with formating and making the colors green for positive and red for negative.


## Conlusion

### Advantages 
Some of the advantages for refactoring code would be that its alot more easier to add new feautures and then make sure the code is good

### Disadvantages 
The bugs included with refactoring a module is way too way. i mustve spent three times more on fixing bugs then i spent creating the code

### Orginal vs Refactored
Orginal has alot more information thrown at you rather then refactored. While refactored is alot more simpler and more "elegant" i would say.

### Summary
This showed how many bugs you can run into with vba in excel. Also showed alot of interesting codes to make excel work differently. Making buttons to run info faster has never been easier!

