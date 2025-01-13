# DataQueryFetch Pro!

This project is a Streamlit application that allows users to execute SQL queries on different Oracle databases and download the results as an Excel file. The application supports large datasets over 10 million/1cr rows within minimum time by splitting them into multiple sheets if necessary.

## Features

- Execute SQL queries on multiple Oracle databases.
- Download query results as an Excel file.
- Handles large datasets by splitting them into multiple sheets.
- Progress bar to show the execution status.

## Installation

To run this application, you need to have Python installed along with the following packages:

- streamlit==1.40.1
- pandas==2.2.3
- oracledb==1.4.2
- XlsxWriter==3.2.0

## Here are some reasons why the .xlsx format might be preferred over CSV:

#### Multiple Sheets: 
The .xlsx format supports multiple sheets within a single file, which is useful for large datasets that need to be split into chunks.

#### Formatting: 
.xlsx files can include cell formatting, such as font styles, colors, and borders, which can make the data easier to read and interpret.

#### Formulas and Functions: 
.xlsx files can contain Excel formulas and functions, allowing for more complex data manipulation directly within the file.

#### Data Types: 
.xlsx files can store different data types (e.g., dates, numbers, text) more accurately than CSV files, which treat all data as plain text.

#### Error Handling: 
The .xlsx format can handle special values like NaN (Not a Number) and inf (infinity) more gracefully.