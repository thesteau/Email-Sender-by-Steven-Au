# Email-Sender-by-Steven-Au
Automatically send bulk emails based on data elements for both the Sender and Recipient(s) using data files.

## Requirements
```
pandas 1.3.2 (Due to read_excel compatibility)
Python >=3.7.x
```

## Getting Started
Two files are required, they can be any file extension.  
*Currently, only CSV and Excel .xlsx are supported.*  
## Important
Please use the provided headers of the CSV and Excel files as guidance.

CSV File:
* Will only read and process the second row. 
* The first row, header, will be ignored.  

Excel File:
* Will only read and process from the second row onwards.
* The first row, header, will be ignored.
* The **sheet name** must be provided when prompted.

*Full file paths are intended to be used for both inputs and attachment referencing.*

### Data File Structure

**Note**: Excel files have an additional requirement of entering the sheet name (tab name).
```
**CSV** - contains the service, port, your login email, your password, and signature only as a HTML.
- All one needs to do to swap services is to modify the CSV file if the data in the Excel file remains the same.

**Excel** - Contains the To addresses, CC addresses, Subject, Body, and Attachments. Multiple *TO, CC, or Attachments* are all read seperately by a ***COMMA (",")***.
```

### Procedure

1. Launch program - import requested Excel file path with elements containing the following values: To_email, CC_email, Email_subject, Email_body, and Email_attachments

2. Choice: Automatic or Manual

3.   
    a. If manual, then the program will open up Google Chrome and create email drafts
    b. If automatic, then the program will request for a CSV file path with the following values: 
SMTP_server_service (Such as smtp.google.com), Port_number, Your_email, Your_password, Your_signature

The password is optional depending on your SMTP server. If omitted, the program will attempt to log in without your password/App password.
Your signature must be saved in a .HTML format


## Port Configurations
The encryption used in this program is TLS and ***not*** SSL 

So the Python code uses the following:
```python
>>> server = smtplib.SMTP(service, port)
```
Otherwise, SSL is simply the following:
```python
>>> server = smtplib.SMTP_SSL(service, port)
```

Where service and port are your smtp service and port used.

## Reference - 
*Only for the Sender file*
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
## Credits
*Code citations are available in-line.*

## Author
Steven Au

## License
See the LICENSE.md for details