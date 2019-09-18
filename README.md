# Email-Sender-by-Steven-Au-Python-3.7
! Run this program accordingly - Run_This_ Email_Sender_Program.py

Procedure:
1. Launch program - execute import of the Excel file elements via PANDAS module
2. Choice of manual or automatic - recommended automatic
3a. If manual, then the assumption is Google Chrome with Drafts
3b. If automatic, then the program will request for a CSV file path to input the following: 
SMTP server service (Such as smtp.google.com), Port (Number only), Your email, Your pass *(* Password Optional - depends on the SMTP service. If omitted, the program will attempt to log in without your password/App password, again varies by service such as Google requesting for an App password)

Please use the provided headers of the CSV and Excel files as guidance.
CSV File - Will only read and process the second row. The first row, header, will be ignored.
Excel File - Will only read and process from the second row onwards. The first row, header, will be ignored.
The program will also need the exact Excel File Sheet name you are reading from. The intent is to have the Sheet read as the 4th sheet.

Full file paths are intended to be used for both inputs and Excel attachment referencing.

External Modules to install:
Pandas (To read both Excel and CSV)
xlrd
pyautogui

Program was written by Steven Au
