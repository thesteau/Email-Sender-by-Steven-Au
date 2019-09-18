# Email-Sender-by-Steven-Au-Python-3.7
! Run this program accordingly - Run_This_ Email_Sender_Program.py

# Motivation
To have a simplified way to send bulk emails via elements stored in a CSV file and an Excel file. 

*(An Excel file is used to create a custom sheet of values so you can have a referenced sandbox sheet)*


# Files
**CSV** - contains the service, port, your login email, your pass, and signature only as a HTML.
- All one needs to do to swap services is to modify the CSV file if the data in the Excel file remains the same.

**Excel** - Contains the To addresses, CC addresses, Subject, Body, and Attachments all read seperately by a ***COMMA (",") only***, applicable to only the *TO, CC, or Attachments* fields.

# Procedure:

#1. Launch program - execute import of the Excel file elements via PANDAS module

#2. Choice of manual or automatic - recommended automatic

#3a. If manual, then the assumption is Google Chrome with Drafts

#3b. If automatic, then the program will request for a CSV file path to input the following: 
SMTP server service (Such as smtp.google.com), Port (Number only), Your email, Your pass *(* Password Optional - depends on the SMTP service. If omitted, the program will attempt to log in without your password/App password, again varies by service such as Google requesting for an App password)

# Port Configurations
The encryption used is TLS and not the SSL variant. 

So the Python code uses the following:
```python
>>> server = smtplib.SMTP(service, port)
```
Otherwise, SSL is simply the following:
```python
>>> server = smtplib.SMTP_SSL(service, port)
```

Where service and port are your smtp service and port used.

# Important
Please use the provided headers of the CSV and Excel files as guidance.
CSV File - Will only read and process the second row. The first row, header, will be ignored.
Excel File - Will only read and process from the second row onwards. The first row, header, will be ignored.
The program will also need the exact Excel File Sheet name you are reading from. *The intent is to read your specified Excel sheet.*


*Full file paths are intended to be used for both inputs and Excel attachment referencing.*


# External Modules to install:
* Pandas (To read both Excel and CSV)
* xlrd
* pyautogui



# Reference - *for the CSV file*
>**Port** - *587*

>**Outlook** - 
*smtp-mail.outlook.com*

>**Yahoo** -
*smtp.mail.yahoo.com*

>**Google** -
*smtp.gmail.com*

>**Hotmail** -
*smtp.live.com*

https://support.google.com/a/answer/176600?hl=en

---
# Credits
*Anything sourced are referenced in the script due to the ever-changing nature of keeping or removing the code*


---
*Scripts in this package were written by Steven Au*
