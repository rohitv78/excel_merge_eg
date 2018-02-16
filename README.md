# Merging Microsoft Excel workbooks

This repository contains an example of how to merge multiple Microsoft Excel workbooks into one workbook. The main reason for this was mainly for scientific manuscript submission where supplemental tables are often merged into a single table. This allows you to have your supplemental tables in different workbooks allowing you to edit each one independently. Then merge it into workbook with all your tables when you are ready.

## How to Use

> The following instructions have been tested on Microsoft Excel for Mac 2011 (Version 14.5.0).

1. Edit the list-of-tables.txt to add the excel workbooks. Each line should be the relative path to the excel workbook you want to merge in.
1. Prepare the individual excel workbooks. **One worksheet per workbook**. 
1. Open the `final-excel.xls` file. Then go `Tools -> Macro`. Choose GetSheets then click `Run`. This will then merge the worksheets from `workbook1.xlsx` and `workbook2.xlsx` and rename them 1 and 2.

## How it Works

The merging is done through visual basic macro in the `final-excel.xls` file. One can see and edit the code by choosing `Tools -> Macro` and then choosing to `Edit` the GetSheets macro. It uses the list-of-tables.txt to figure out which excel workbooks to merge in.
