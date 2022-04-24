# Stock-Analysis
Analyzing stocks market financial data by using VBA in Excel.

# Overview of Project: 
In this project, we are helping Steve, who’s just graduated with his finance degree. His parents will be his first client and are passionate about green energy stocks. Steve asked us to help him evaluate the all green energy stocks and DAQ stocks that his parents had planned to invest. 
Steve loved the workbook we prepared for him. He wants us to expand the dataset to include the entire stock market over the last few years. Although our code works well for dozen stocks, it might not work as well for thousands of stocks. Also, it may take a long time to execute. 
This project aims to edit or refactor the code to loop through all the data one time to collect the same information that we have already prepared for him in this project. Then, we determined whether refactoring code successfully made the VBA script for optimization. Using code to automate analyses allows Steve to reuse it with any stocks and reduces the chance of accidents, errors, and runtime.
# Results: 
## Stock performance between 2017 and 2018.
First, we created a VBA script for 2017 stocks where all the stocks had positive return except TERP. Moreover, DQ return percentage was at the third place where Steve parents are interested to invest.

<img width="235" alt="2017" src="https://user-images.githubusercontent.com/92646311/164845511-041a0b50-f6ee-4411-9d12-98aa6736c86f.png">

Next, we created the VBA script for 2018 to further explore the stocks to see the difference in Return and Total Daily Volume between the two years. We can observe that RUN stocks had a significantly high Total Daily Volume and return in 2018 compared to 2017, which were still positive. The ENPH, Total Daily Volume increased, and the Return percentage slightly dropped, still had positive results.

<img width="239" alt="2018" src="https://user-images.githubusercontent.com/92646311/164845563-89e59dfa-ef70-410a-8298-91163abbba66.png">

## Execution times of the original script and the refactored script.
**Orignal Script run Time for 2017**

<img width="244" alt="PopUP_2017" src="https://user-images.githubusercontent.com/92646311/164847825-8ea9cb9f-002a-4974-81a0-6897c780b833.png">

**Orignal Script run Time for 2018**

<img width="247" alt="PopUp_2018" src="https://user-images.githubusercontent.com/92646311/164851902-b742cb9a-ff32-445a-b53f-a259a1e09153.png">

**Run time for Refactored code 2017**

<img width="252" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/92646311/164853298-de<img width="245" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/92646311/164984818-9f59f276-af8b-4f8d-9dac-4540e998c2bf.png">
e62014-089e-4142-a6f8-f69cf6ad8a51.png">

**Run time for Refactored code 2018**

![Uploading VBA_Challenge_2018.png…]()

# Summary 

1.	What are the advantages or disadvantages of refactoring code?

Adventages

•	The best advantage of refactoring code is that it is executed more efficiently and makes programs run faster. 

•	It makes the code clean, short, and more organized than the long lines of code. 

•	It improves the code quality, decreases the chance of errors, and reduces the time needed to run analyses, especially if they need to be done repeatedly.

Disadvantage

•	It is intricate and problematic to deconstruct the code. Editing is complicated. 

•	It is not easy and time-Consuming. 

2.	How do these pros and cons apply to refactoring the original VBA script?

It optimizes the performance of the script and reduces the run time. It helps remove the complexions of the code and enhance the code's reliability. Code refactoring also helps remove vulnerabilities of the system. It removes bugs. There are possible chances that it goes wrong due to the complexity of the code, and you will have to spend more time solving the problem.


