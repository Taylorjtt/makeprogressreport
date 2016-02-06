
makeprogress report is a python script that automates making a new progress report.

Each week you run the script and it makes a new progress report with the proper naming
convention and sets the start date to the previous saturday.

setup steps
-----------
1. Get Python > 3.0 from https://www.python.org/downloads/
2.Once you have python installed open a cmd prompt and run
  python -m pip install openpyxl
  This installs the library to edit excel worksheets
3. extract the Progress_Template.xlsx and makeprogressreport.py files to the folder where
   you save  your progress reports
4. run makeprogressreport.py