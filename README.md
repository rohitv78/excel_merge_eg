# merge-excel-workbooks

## How to Use

> The following instructions have been tested on Microsoft Excel for Mac 2011 (Version 14.5.0).

1. Edit the list-of-tables.txt to add the excel workbooks. Each line should be the relative path to the excel workbook you want to merge in.
1. Prepare the individual excel workbooks. **One worksheet per workbook**.
1. Open the `final-excel.xls` file. Then go `Tools -> Macro`. Choose GetSheets then click `Run`. This will then merge the worksheets from `workbook1.xlsx` and `workbook2.xlsx` and rename them 1 and 2. 
