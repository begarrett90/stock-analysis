# stock-analysis

## Overview of Project

### Purpose
  The purpose of this project was to be able to analyze an entire dataset of the stock market over the last few years with the click of a button. Original code was refactored to loop through all the data at one time to collect the same information instead of using the original nested for loops to see if it made the code more efficient. 
  
  The refactored code should allow a user to put a year into an input box and give formatted results that are easy to read for the user. The original code required seperate buttons to run the analysis and format the results. The original code and the refactored code used a timer to see how long the analysis took allowing us to determine which code is more efficient. 

## Results

### Stock Performance 2017 vs 2018 

  In 2017 most stocks performed well, with all but one having a positive return. In 2018 however most stocks had a negative return except for ENPH and RUN. ENPH and RUN had positive return in both 2017 and 2018. In 2017 and 2018 ENPH had a total daily volume over 200,00,000 and 600,000,000 and a return of 129.5% and 81.9%. In 2017 and 2018 RUN had a total daily volume over 200,00,000 and 500,000,000 and a return of 5.5% and 84%. The biggest lost on return in 2018 was seen in DQ and JKS each with over-60% return. According to this analysis the safer stocks in invest in would be ENPH and RUN. 

### Original Code vs Refactored Code Peformance

  The original code perfomed based on nested for loops and required use of a second code/button in order to properly format the results. The refactored code used one for loop and the formatting for the results was included in the refactored code. This made the refactored code easier to use, as you only had to run one code to have properly formatted results vs. the original code which required you to run two seperate codes to get the results and then properly format. Below are pictures of the code that show the original codes use of nested for loops and the refactored code use of one loop. Not shown in the picture is the formatting portion of the code.
  
  **The original code:**
  
  <img width="522" alt="Original Code" src="https://user-images.githubusercontent.com/105942622/173242658-8933b60b-1760-472e-a512-38b8359edfc7.png">

  **The refactored code:** 
  
<img width="530" alt="Refactored Code" src="https://user-images.githubusercontent.com/105942622/173242875-592c1797-59fe-48b4-8f2a-85d660139386.png">

The next metric used to measure performance was the amount of time that it took to run the code. 

The original code in the year 2017 took 0.5507813 seconds and in 2018 it took 0.5625 seconds.

<img width="226" alt="VBA 2017 original code" src="https://user-images.githubusercontent.com/105942622/173243402-cb34c047-468d-4e00-8512-9b5b97a7e6ab.png">
<img width="215" alt="VBA 2018 original code" src="https://user-images.githubusercontent.com/105942622/173243440-75923d23-f8b3-4334-a11a-edc69261f1dc.png">

 The refactored code ran in 0.0625 seconds in 2017 and 0.0703125 seconds in 2018. 
 
<img width="222" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/105942622/173243097-739da750-95fb-442a-8686-2de4825b0434.png">
<img width="232" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/105942622/173243102-2892a473-097e-40df-af05-28e418a7e706.png">

The refactored code decreased the time to run the code by 88% for 2017 and 87.5% in 2018

## Summary

Refactoring the code allowed for a more efficient code which ran in fewer steps, used less memory and improved the logic of the code which will make it easier for future readers of the code. 

### Advantages & Disadvantage of refactoring Code

**Advantages**
1. Fewer steps to run the code, eliminating duplicated routines and unnecessary loops
2. Less memory used 
3. Improved logic and easier to read in the future
4. Refactoring code can help eliminate bugs or remove unnecessary or faulty code

**Disadvantages**
1. Time consuming
2. Can be harder to perform on larger code and increase risk of person refactoring causing bugs or faults in the code
3. Takes work/resources

### How pros and cons apply to refactoring the original VBA script
Refactoring the VBA stock analysis code reduced the loops required to run the code resulting in a more efficient code with decreased run time. The refactored code requires less memory and processing power to run the code. The refactored code also is more logical and easier to read. The cons of refactoring the code is that it took time to refactor the code and that in order to determine if the refactoring was successful required testing and running the code multiple times, fixing bugs along the way, and comparing run times to the original code. 
