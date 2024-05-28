# Data Imputation
Fill in missing data using python. 

The purpose of this script is to fill missing values in an Excel sheet and update the "STATE" column based on the most frequent values associated with each "CITY". The script performs the following tasks:

**Load Excel Workbook:** The script uses openpyxl to load an Excel workbook and select "Sheet1".

**Define Helper Functions:**
**reve():** This function fills a cell in column "STATE" with a value from the previous row if the value in column "CITY" matches the previous row's value, or it calls forwa() if not.

**forwa():** This functions iterates through the "CITY" column in forward direction and checks for matching value. If the "CITY" is same then it checks for a value in "STATE", "COUNTY", and "PROVINCE" column respectively. If a value is found then it is copied into the "STATE" column of the original row.
  
For example:

CITY    PROVINCE  COUNTY      STATE
Mumbai
Mumbai            Maharashtra
        
Consider this data set. The STATE, COUNTY, and PROVINCE values are empty. So, the forwa() function wil check whether the next column has the same City or not and if it does then it will detect the value in COUNTY column and paste it in the STATE column of the first row of this sample dataset.


**Iterate Through Rows:** The script iterates through rows, checking if cells in column "E" are empty. If so, it attempts to fill them with values from columns "C" or "D". If these are also empty, it calls reve() to handle the missing value.

**Save the Workbook:** The updated Excel workbook is saved to a new file.

**Process with Pandas** The script then uses pandas to read the saved workbook, convert all entries in the "STATE" column to strings, and find the most frequent province for each city.

**Update the "STATE" Column:** It updates the "STATE" column in the DataFrame based on the most frequent province for each city.

**Save the DataFrame:** Finally, the script saves the updated DataFrame in a new Excel file.

Overall, this script aims to clean and standardize the data by filling in missing values and ensuring consistency in the "STATE" column based on city data.
