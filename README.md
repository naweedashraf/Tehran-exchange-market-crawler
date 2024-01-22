This project contains 3 files:
1- fetchData.py
2- XLSX2CSV.py
3- analyze.py

Following libraries must be installed to run the files correctly:
argparse==1.1
csv==1
jdatetime==4.1.1
logging==0.5.1.2
openpyxl==3.1.2
requests==2.31.0


fetchData.py
=======================
Use this file to acquire market watch excel files of Tehran stock exchange market.
In order to run this file you must enter start and end date in the format of YYYY-MM-DD(Jalali Date).

example: $ python fetchData.py 1401-12-27 1402-01-15

XLSX2CSV.py
=======================
Use this file to convert excel files to csv files.
In order to run this file you must enter source folder, which contains excel files, and a boolean argument to wheter remove excel files after converting or not.

example: $ python XLSX2CSV.py stage True

analyze.py
=======================
Use this file to analyze the data in csv files.
Result of the analyzes will be saved in info.log

run: $ python analyze.py