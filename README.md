# Email-Sender-by-Steven-Au-Python-3.7
! To run this program accordingly - Run_This_ Email_Sender_Program.py

# Motivation
To have a simplified way to send bulk emails via elements stored in a CSV file and an Excel file. 

*(An Excel file is used to create a custom sheet of values so you can have a referenced sandbox sheet)*


# Files
**CSV** - contains the service, port, your login email, your password, and signature only as a HTML.
- All one needs to do to swap services is to modify the CSV file if the data in the Excel file remains the same.

**Excel** - Contains the To addresses, CC addresses, Subject, Body, and Attachments. Multiple *TO, CC, or Attachments* are all read seperately by a ***COMMA (",") only***.

# Procedure

#1. Launch program - import requested Excel file path with elements containing the following values: To_email, CC_email, Email_subject, Email_body, and Email_attachments

#2. Choice: Automatic or Manual

#3a. If manual, then the program will open up Google Chrome and create email drafts

#3b. If automatic, then the program will request for a CSV file path with the following values: 
SMTP_server_service (Such as smtp.google.com), Port_number, Your_email, Your_password, Your_signature

Passord is optional depending on your SMTP server. If omitted, the program will attempt to log in without your password/App password.
Your signature must be saved as a .HTML format

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
The program will also need the exact Excel File Sheet name you are reading from. 


*Full file paths are intended to be used for both inputs and Excel attachment referencing.*


# External Modules to Install
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
