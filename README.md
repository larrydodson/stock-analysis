# stock-analysis
Module 2 UTMCC_DataViz, VBA_Challenge
---

## Contents

### 
  1. Overview of Project, and Background
  2. Results 
  3. Summary 

---

## 1. Overview of Project
### **Purpose**

  The primary purpose of the project is to utilize VBA (Visual Basic for Applications) coding script language to prepare summary financial charts with data on selected equity stocks as is of interest by our associate, Steve. The resulting charts generated from the VBA code were then used for comparison analysis of the various competitve stocks chosen by Steve, then to be used for review of it's perfomrance in the specified years, and also for financial planning for future investments. 
  
  An addtional purpose of the exercise was to create and develop proficiency with VBA by creating macros within Excel, utilizing For-loops, If-then statements, formatting the outputs within Excel sheets for readability, creating interactivity for ease-of-use, and to understand the value of refactoring the original code for increasing performance and reliability. 

### **Background**

  Steve requested assistance to evaluate green-energy stocks, based primarily on the investment of his parents into a single company's stock, named DAQO (ticker DQ). The single company investment is risky, and Steve wishes to evaluate several others stocks along-side DAQO so that he may discuss with his parents about diversifying investments into other companies to create a general portfolio. Steve has obtained a data set of the years 2017 and 2018 of twelve different company stocks he wishes to evaluate. Data analysis and visualizations on these stocks' financial performance were then completed.
  
  From this process the VBA macro was created also as a portable financial tool that will allow Steve to quickly run the performance chart now and in the future on new data, as years and companies to be evaluated will change. The below information is a result and summary of the using the VBA tools. 

---
## 2. Results 

In Figure-1, below, we see an example of the data as provided for DQ and is representative for all company stocks in the master table. For both years 2017 and 2018 the avaialble data was for the identifying ticker symblol, and daily information on the stock's price for Open, the High, the Low, the daily Close price, and for the daily stock shares traded volume. The data was available on twelve different companies. (Note: the data column for "Adj Close" was not used in this evaluation.) 

**Figure-1** "Data Table, partial for DQ as example"
![DQchart_raw_data_2018.PNG](https://github.com/larrydodson/stock-analysis/blob/master/resources/DQchart_raw_data_2018.PNG)


Within Excel creating imbedded VBA macro subroutines, the following information was produced and presented.  
Please see Figure-2 below. 

| Year | 2017 | 2018 |
| ---:         |     :---:      |          :---: |
| Data Run, using original code script | ![orig_code_runtime_2017_848.PNG](https://github.com/larrydodson/stock-analysis/blob/master/resources/orig_code_runtime_2017_848.PNG) | ![orig_code_runtime_2018_832.PNG](https://github.com/larrydodson/stock-analysis/blob/master/resources/orig_code_runtime_2018_832.PNG) |

**Figure-2**


.

**Stock Share Performance**
- In analysis of the stocks' performance in 2017, DQ's stock outperformed all others with a 199.4% return for the year, and only one company with a negative performance, TERP. 
- For perfromance in 2018, there were very different results, with only two stocks with positive returns, ENPH and RUN. All other stocks had negative returns. DQ was in the negative return group at -62.6%. 


.

**Code Execution Performance**  (measuring execution times of the original script and the refactored script) 
- After verifying code accuracy, the original version of the code was evaluated and then refactored. Please see below for more information on the refactoring process.  
- See Figure-3 below. 

| Year | 2017 | 2018 |
| ---:         |     :---:      |          :---: |
| Original Code - prior to Factoring | 0.848 s  | 0.832 s |
| New Code - after Factoring    | ![VBA_Challenge_2017 184.PNG](https://github.com/larrydodson/stock-analysis/blob/master/resources/VBA_Challenge_2017%20184.PNG) | ![VBA_Challenge_2018 195.PNGe](https://github.com/larrydodson/stock-analysis/blob/master/resources/VBA_Challenge_2018%20195.PNG) |
| .    | ![VBA_Challenge_2017.PNG](https://github.com/larrydodson/stock-analysis/blob/master/resources/VBA_Challenge_2017.PNG) | ![VBA_Challenge_2018.PNG](https://github.com/larrydodson/stock-analysis/blob/master/resources/VBA_Challenge_2018.PNG) |
| Improvement in run time execution   | 78% | 77% |

**Figure-3**

- Through Factoring, improvement in execution times was realized for code performance by 78% for 2017, and 77% for 2018. 
- The Factoring process included editing the original code script from processing of the major variables of share prices and volumes from a nested For-loop, to running various For-loops, that were not nested but instead using an incremental Ticker Index. 

.

---
## 3. Summary
  
- VBA as a stock planning financial tool:  Steve was very happy with the results and use of the tool. It gave him and his parents the capability and visibility to evaluate the stocks of interest, and enabled them to create a stock portfolio of investments.    
  
  
  **Refactoring Code** 
  In general refactoring code is the editing and optimization to clean it, reorganize and improve the logic order. It comes with several advantages, and is usually considered good professional practice. 
  
  This project's refactored code solution included a loop through all the data one time, versus the multiple loops and nested loop in the original code that generated the same output information. The new refactored code was more efficient and the result was a 77% improvement in code exectution time.


- **Refactoring Advantages**
  As a list, refactoring typically will lead to performance improvements due to efficiency, lower use of memory and architecture stacks, increased interpretation or compiling efficiency, overall cost reduction to a system, code readability, reliabilty and quality, agility and portability (to other projects and platforms and languages), and increased code execution speed.
 
 - **Refactoring Disadvantages**
  Reasons to not refactor code may be due to delivery deadlines, where there is insufficient test time available to complete refactoring, and the original code is performing adequately, And, where costs, time and people resources to perform and test the refactored product are not a priority, where the resources are needed for other higher priority projects, and codes slated for refactoring are not high critical importance, it may not be justifiable to refactor. Also, there may not be a need to refactor code if it is known to have a short useful life, and again where time and resources are not available to improve it's run time or quality.

.end 
