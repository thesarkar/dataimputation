# DataImputation
Fill in missing data using python. 

The purpose of this script is to fill missing values in an Excel sheet and update the "Province" column based on the most frequent values associated with each "CITY". The script performs the following tasks:

**Load Excel Workbook:** The script uses openpyxl to load an Excel workbook and select "Sheet1".
**Define Helper Functions:**
**reve():** This function fills a cell in column "E" with a value from the previous row if the value in column "B" matches the previous row's value, or it calls forwa() if not.
**forwa():** This function searches forward from the current row to find and copy a value from columns "D" or "C" to column "E" when the value in column "B" matches.
**Iterate Through Rows:** The script iterates through rows, checking if cells in column "E" are empty. If so, it attempts to fill them with values from columns "C" or "D". If these are also empty, it calls reve() to handle the missing value.
**Save the Workbook:** The updated Excel workbook is saved to a new file.
**Process with Pandas** The script then uses pandas to read the saved workbook, convert all entries in the "PROVINCE" column to strings, and find the most frequent province for each city.
**Update the "PROVINCE" Column:** It updates the "PROVINCE" column in the DataFrame based on the most frequent province for each city.
**Save the DataFrame:** Finally, the script saves the updated DataFrame back to the Excel file.

Overall, this script aims to clean and standardize the data by filling in missing values and ensuring consistency in the "STATE" column based on city data.
