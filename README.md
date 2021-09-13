# stock-analysis

## Overview: VBA Module 2 Challenge

### Purpose
The purpose of this project is to refactor (edit) the code provided in the challenge in order to make the script
run faster. 

The data set shows stock information for two years, 2017 & 2018. Our goal is to improve the effiency
of the code presented & determine whether there is opportunity in future investing in these stocks or not. 

Our results will then be presented to Steve. With the new refactored script, Steve will be able to analyze the 
entire dataset and make an informed decision for his parents regarding which stocks to pursue invesment in (if any).
___
## Data Background

> Steve loves the workbook you prepared for him. At the click of a button, he can analyze an entire dataset. Now, to do a little more research for his parents, he wants to expand the dataset to include the entire stock market over the last few years. Although your code works well for a dozen stocks, it might not work as well for thousands of stocks. And if it does, it may take a long time to execute.

> In this challenge, you’ll edit, or refactor, the Module 2 solution code to loop through all the data one time in order to collect the same information that you did in this module. Then, you’ll determine whether refactoring your code successfully made the VBA script run faster. Finally, you’ll present a written analysis that explains your findings.

> Refactoring is a key part of the coding process. When refactoring code, you aren’t adding new functionality; you just want to make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read. Refactoring is common on the job because first attempts at code won’t always be the best way to accomplish a task. Sometimes, refactoring someone else’s code will be your entry point to working with the existing code at a job.

### Challenge 
> Use your knowledge of VBA and the starter code provided in this Challenge to refactor the VBA Module 2 Script so you loop through the data one time and collect all of the information. 
> 
> Your refactored code should run faster than it did in the module.
___

## Results
> Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.

1. The `tickerIndex` is set equal to zero before looping over the rows.

    ![image](https://user-images.githubusercontent.com/89520192/133012396-2d66e7db-5fb1-4807-b7eb-a600ca02c7e3.png)

2. Arrays are created for `tickers`, `tickerVolumes`, `tickerStartingPrices`, `tickerEndingPrices`

    ![image](https://user-images.githubusercontent.com/89520192/133012545-244c9292-3046-4d02-b2f5-c54c1599eb73.png)

    ![image](https://user-images.githubusercontent.com/89520192/133012561-b6868281-742c-467e-b120-848bb45e820e.png)

3. The `tickerIndex` is used to access the stock ticker index for the `tickers`, `tickerVolumes`, `tickerStartingPrices`, and `tickerEndingPrices` arrays

    ![image](https://user-images.githubusercontent.com/89520192/133012677-c220b604-e1d3-41f5-98d8-99ad2d9f6d46.png)

4. The script loops through stock data, reading and storing all of the following values from each row: `tickers`, `tickerVolumes`, `tickerStartingPrices`, and `tickerEndingPrices`

    ![image](https://user-images.githubusercontent.com/89520192/133012841-7fea6f98-0289-4326-a310-ff4fd13123f6.png)

5. The outputs for the 2017 and 2018 stock analyses in the `VBA-Challenge.xlsm` workbook match the outputs from the AllStockAnalysis in the module
6. The pop-up messages showing the elapsed run time for the script are visible an saved as `VBA_Challange_2017.png` and `VBA_Challenge_2018.png`

      **1.1 - VBA_Challenge 2017 Results** ✅Faster✅
      
      ![image](https://user-images.githubusercontent.com/89520192/133014424-0c65b872-b3d8-4717-9430-8dae420ee06a.png)
      
    **1.2 - Module 2017 AllStockAnalysis** ⭕Slower⭕

      ![image](https://user-images.githubusercontent.com/89520192/133013520-3c3ca582-42a9-4177-9057-b5d4560a5e62.png)     
      
      ![image](https://user-images.githubusercontent.com/89520192/133013728-92b0a6d8-c91e-488c-bdee-c698e5eddd41.png)

    **2.1 - VBA_Challenge 2018 Results** ✅Faster✅
    
      ![image](https://user-images.githubusercontent.com/89520192/133014138-bfd047ef-55ed-458a-9f0c-8de771a9cb2d.png)
   
    **2.2 - Module 2018 AllStockAnalysis** ⭕Slower⭕

      ![image](https://user-images.githubusercontent.com/89520192/133013598-32405a80-39dd-4465-be52-f20accda1dbb.png)

      ![image](https://user-images.githubusercontent.com/89520192/133013674-b25baa28-4ae5-4005-bb0c-81206bd5a8b3.png)
      
* **Module 2017/2018 AllStockAnalysis screenshots match the VBA_Challenge 2017/2018 results.**
* **The VBA_Challange results have significantly reduced elapsed run times.**
* **The refactored script was successful in reducing run-time for this program, so Steve is able to get his data accurately and faster than before!**
___

## Summary Statement

**1. What are the advantages or disadvantages of refactoring code?**

**Advantages:**

* Improved organization of script.
* Quicker exectuion of code with faster run-time results.
* Easy to follow, readable logic. This may help for debugging purposes, or to allow someone not familiar with the code understand it easier. 

**Disadvantages:**

* Can be too costly.
* Can be too time consuming if the original code is too large or function specific. 
* Failure to understand original function of code can cause refactoring to output different results from those originally intended.

**2. How do these pros and cons apply to refactoring the original VBA script?**

**Pros:**

* Macro runs much faster.

   **As can be seen below, the macro for 2017 VBA_Challenge runs in 0.0859375 seconds while the code from the original Module solution runs in 0.7109375 seconds.**

    ![image](https://user-images.githubusercontent.com/89520192/133016427-d3e4fcb2-5623-4885-b63e-53d12eda5c63.png)

    vs.

    ![image](https://user-images.githubusercontent.com/89520192/133016478-b2c990e3-f772-41a5-9aa9-0e6b5c48df05.png)

**Cons:**

* Understanding VBA and how to code in general is a little difficult for most people. On top of that, refinement of code requires special attention and knowledge of syntax and logic. The refactoring of the original VBA script took time and lots of debugging until it ran smoothly with no errors.
