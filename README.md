# stock-analysis
Module 2 UTMCC_DataViz, VBA_Challenge
---

## Contents

### 
  1. Project Overview
  2. Results 
  3. Summary 

---

## 1. Overview of Project
### **Purpose**

  The primary purpose of the project is to utilize VBA (Visual Basic for Applications) coding script language to prepare summary financial charts with data on selected equity stocks as is of interest by our associate, Steve. The resulting charts generated from the VBA code was then used for comparison analysis of the various competitve stocks, then to be used for review of it's perfomrance in the specified years, and for financial planning for future investments. 
  Addtional purpose of the exercise was to create and develop proficiency with VBA by creating macros within Excel, utilizing For-loops, If-then statements, formatting the outputs within Excel sheets for readability, creating interactivity for ease-of-use, and to understand the value of refactoring the original code for increasing performance and reliability. 

### **Background**

  Steve requested assistance to evaluate green-energy stocks, based primarily on the investments of his parents into a single company's stock, named DAQO (ticker DQ). The single company investment is risky, and Steve wishes to evaluate several others stocks along-side DAQO so that he may discuss with his parents about diversifying investments into other companies to create a portfolio. Steve has obtained a data set of the years 2017 and 2018 of twelve difference compay stocks he wishes to evaluate, and data analysis and visualizations on how these stocks financially performed were completed.
  
  From this process the VBA macro was created also as a portable financial tool that will allow Steve to quickly run the performance chart now and in the future on new data, as the years and companies to be evaluated will change. The below information is a result and summary of the using the VBA tools. 

---
## 2. Results 

In Figure-1, below, we see an example of the data as provided for DQ and is representative for all company's stock in the master table. For both years 2017 and 2018 the avaialble data was for the identifying ticker symblol, and daily information on the stock's price for Open, the High, the Low, the daily Close price, and for the daily stock shares volume that was traded. The data  was available on twelve different company's stock information in this format. (Note: the data column for "Adj Close" was not used in this evaluation.) 

Figure-1 "Data Table example for DQ"
![DQchart_raw_data_2018.PNG](https://github.com/larrydodson/stock-analysis/blob/master/resources/DQchart_raw_data_2018.PNG)


Within Excel creating imbedded VBA macro subroutines, the following information was produced and presented.  
Please see Figure-2 below. 

| Year | 2017 | 2018 |
| ---:         |     :---:      |          :---: |
| Data Run, using original code script | ![orig_code_runtime_2017_848.PNG](https://github.com/larrydodson/stock-analysis/blob/master/resources/orig_code_runtime_2017_848.PNG) | ![orig_code_runtime_2018_832.PNG](https://github.com/larrydodson/stock-analysis/blob/master/resources/orig_code_runtime_2018_832.PNG) |

Figure-2


.

**Stock Share Performance**
- In looking at the the stocks' performance in 2017, DQ's stock outperformaed all others with a 199.4% return for the year, and there was only one company that had negative performance, TERP. 
- For perfromance in 2018, it was a much different results overall, with only two stocks with positive returns, ENPH and RUN, and all other with negative return. DQ was in the negative return group with a return of -62.6%. 


.

**Code Execution Performance**  execution times of the original script and the refactored script
- After verifying code accuracy, the original version of the code was evaluated and then refactored. In the Summary below, please see description and purpose of this process. 
- See Figure-3 below. 

| Year | 2017 | 2018 |
| ---:         |     :---:      |          :---: |
| Original Code - prior to Factoring | 0.848 s  | 0.832 s |
| New Code - after Factoring    | ![VBA_Challenge_2017 184.PNG](https://github.com/larrydodson/stock-analysis/blob/master/resources/VBA_Challenge_2017%20184.PNG) | ![VBA_Challenge_2018 195.PNGe](https://github.com/larrydodson/stock-analysis/blob/master/resources/VBA_Challenge_2018%20195.PNG) |
| .    | ![VBA_Challenge_2017.PNG](https://github.com/larrydodson/stock-analysis/blob/master/resources/VBA_Challenge_2017.PNG) | ![VBA_Challenge_2018.PNG](https://github.com/larrydodson/stock-analysis/blob/master/resources/VBA_Challenge_2018.PNG) |
| Improvement in run time execution   | 78% | 77% |

Figure-3




.


![image_name](path/to/image_name.png)


![image_name](path/to/image_name.png)


.

---
## 3. Summary
  
- Refactoring code - 
detailed statement on the advantages and disadvantages of refactoring code in general
 
 - Advantages 
 
 
 - Disadvantages
  
  
- How do these pros and cons apply to refactoring the original VBA script?

a detailed statement on the advantages and disadvantages of the original and refactored VBA script 


edit, or refactor, the Module 2 solution code to loop through all the data one time in order to collect the same information that you did in this module. Then, you’ll determine whether refactoring your code successfully made the VBA script run faster.





.end 
