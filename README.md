# FED-Watch-Tool

## Requirements

**Python Libraries:**
1. pdblp
2. blpapi

**xlwings**

* Reference to installation doc: [Link]()
* General help: [Link]()
  
**Anaconda/Jupyter Notebooks**
  
  Refernce to installation: [Link](https://www.anaconda.com/)
  
**Microsoft Visual Studio Code**
  
  Reference to installation: [Link](https://code.visualstudio.com/)

**Bloomberg Terminal**
  
## Description

The tool provides a visual representation of the expected path of policy rate changes. The tool inputs post meeting implied rates (tickerized in bloomberg
eg. USOAFR MMMYYYY) to calculate percent of a single hike (+) or cut (-) at the associated meeting. Subsequently, this percent of a single hike/cut is used to
calculate unconditional probabilities of Federal Open Market Committee (FOMC) meeting outcomes to generate a binary probability tree.
			
![image](https://github.info53.com/storage/user/2781/files/33bc775f-ffa6-4ba1-9af8-4ace792e135a)

Location in shared drive: *Z:\Enterprise Shares\Risk Management\Market_Risk\Fifth Third Securities\Reports\Daily Risk Report 3\Python_codes\fed rate hike probabilities*

## General Framework
																				
![image](https://github.info53.com/storage/user/2781/files/f9ea6ed7-995e-448b-83b3-61f98ae37034)


## Updating & Debugging

The code needs to be updated after every FOMC meeting as the old generic ticker for post meeting implied rate would fall off 
& a new one is added. 

1. In the main function, there is a list of meeting dates - delete the date which fell off & add a new one. *Line 234*

**Code before meeting:**

![image](https://github.info53.com/storage/user/2781/files/51fbb730-14b8-4eb7-80c0-9226e3cbfbeb)

**Code after meeting:**

![image](https://github.info53.com/storage/user/2781/files/97ff1761-84b6-4960-946c-fce4a021a73a)

2. Next, in line 238, edit the argument appropriately to reflect the meeting dates for which the generic ticker is available

**Code before meeting:**

![image](https://github.info53.com/storage/user/2781/files/f76a5957-3189-422c-b652-d2a35e0a2bb8)

**Code after meeting:**

![image](https://github.info53.com/storage/user/2781/files/216b8b07-cca3-44ba-bc26-78e14a968bbc)

3. Printing (location) of output in excel files:

The output of dataframes is pasted in "BBG_Data" tab of the Fed Hike Probabilities tool at cell locations "B3", "B33", "B63" based on the code in lines 259, 261 & 263
as shown in image below.
![image](https://github.info53.com/storage/user/2781/files/119910f9-660d-45ca-b8d3-06b779873380)





