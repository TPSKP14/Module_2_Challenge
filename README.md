# Module_2_Challenge
Module 2 Challenge

Background
Steve loves the workbook you prepared for him. At the click of a button, he can analyze an entire dataset. Now, to do a little more research for his parents, he wants to expand the dataset to include the entire stock market over the last few years. Although your code works well for a dozen stocks, it might not work as well for thousands of stocks. And if it does, it may take a long time to execute.
In this challenge, we will edit, or refactor, the Module 2 solution code to loop through all the data one time in order to collect the same information that we did in this module. Then, we’ll determine whether refactoring the code successfully made the VBA script run faster. Finally, we’ll present a written analysis that explains our findings.


Results 
Before refactoring the code the resulting times to execute were 0.7929688 seconds for 2017 data and .8085938 seconds for 2018 data
After refactoring the code the results were .125 seconds for the year 2017 and .14065 for the year 2018


Conclusion 
Refactoring the code made the scripts significantly faster, approx. six times faster in each instance.


Summary
What are the advantages or disadvantages of refactoring code?

Advantages
•	The code can become more efficient and less ‘clunky’ by taking fewer steps to execute the same goal
•	After refactoring and practicing ‘DRY’ – the code can be easier to read
•	The code uses less memory
•	The code runs faster

Disadvantages
•	The more you change, the more likely you can corrupt or damage the original file and create bugs
•	Difficulties may arise if the original code is not correctly understood
•	Can be an instance of diminishing returns, where the amount of work put in does not justify the results


How do these pros and cons apply to refactoring the original VBA script?
The original VBA scripts run time was reduced dramatically, relatively speaking.  If it were a script that took longer to run we would notice the changes, but in this case it’s difficult to see a significant impact other then that recorded by the program’s timer.  Meaning, human observation does not see much difference between the first and second set of code.  With that being said, a deeper understanding of the code was instilled by the end of the project.  Additionally, the code is easier to follow and is more efficient, if we were to continue building the code out or apply it to a larger data set – then surely we would begin to notice a more significant difference
