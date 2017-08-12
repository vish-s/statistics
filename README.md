# Statistics
A python module to obtain numerous statistical data points for a given data set

Stats.py allows you to obtain Stats - Mean, Min, Max, Std Deviation, Median, 95 Percentil, 99 Percentils - for any "Column"
in a given excel sheet.

This module makes use of the Pandas Library to calculate each stat. 

## Pre Requisites

Required Modules

*Pandas*

## Getting Started


`python stats.py <FILE_NAME.xlsx> <Column_Names>`

Example,

spreadsheet: server_1_data.xlsx

python stats.py server_1_data.xlsx 'Time' 'PING Time' 'DNS Active' 



