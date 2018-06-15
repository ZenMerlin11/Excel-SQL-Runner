# Excel SQL Runner
This is a data extraction utility used to connect to and run a series of user defined SQL Queries on a specified SQL database. The results of each query are stored on sheets created by the user for future analysis. This application is designed for Excel users who want to extract data from SQL databases for further analysis or presentation and need a portable solution that does not require installation.

## Dependencies
This project uses the [Rubberduck](http://rubberduckvba.com/) add-in as a test framework and refactoring tool for the VBA Editor (VBE). The project also requires a reference to the Microsoft ActiveX Data Objects 2.8 Library to run database operations. The macros are designed to be compatible with Excel 2013 and 2016. 

## Conventions
- [Coding Style Guidelines](https://github.com/danwagnerco/vba-style-guide)
- [Docs Style/Tools Used](http://www.naturaldocs.org/)

## Installation
Just download the excel file from the root folder of this repository and modify as needed. Alternatively, the individual .bas files can be downloaded and imported into other spreadsheets as needed.

## Usage
1. Scroll down on the `Connection` worksheet and press the `Toggle Dev Mode` button. Enter the password to enable dev mode. The password is by default a null string, but can be changed by altering the `str_DEV_PASSWORD` constant stored in the `MGlobalConstants` module.
2. Copy and paste queries you want to run in the `Queries` column on the `SQL` worksheet. The Macro will execute all queries contained in consecutive cells starting at the first row below the header (row 3) going down. Execution stops when the first blank cell is encountered.
3. Create and format output worksheets to store your query results. Records will be written to these sheets starting at cell `A6` but can be changed by modifying the `str_DATA_TAB_FIRST_ROW` and `lng_DATA_TAB_FIRST_ROW` values in the `MGlobalConstants` module to the desired values.
4. Add these worksheet names to the cells directly to the right of their corresponding query. Note: If any of the names are misspelled, or otherwise the sheet does not exist, an error will occur.
5. Add `TRUE` or `FALSE` to the next cell to the right to specify whether the output sheet will be made visible after the query is run.
6. Add any find/replace text strings to process on the queries before running in the cells to the right of these. You can add as many string replacements as needed. Find/replace text processing will halt when the first blank cell is encountered along the row.
7. Modify the `Summary` sheet to present any overall data analysis desired.
8. Close developer mode by again pressing the `Toggle Dev Mode` button on the `Connection` worksheet. This will hide all sheets but `Connection`.
9. Enter log in credentials for the database and click `Start Extraction` to begin querying the database. Click `Reset` if it is desired to clear the extracted data.

## Docs
Source code documentation is available [here](https://zenmerlin11.github.io/Excel-SQL-Runner/)
