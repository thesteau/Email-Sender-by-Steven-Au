# Bulk Email Sender by Steven-Au
Automatically send bulk emails based on data elements for both the Sender and Recipient(s) using data files.

## Requirements
```
pandas 1.3.2 (Due to read_excel compatibility)
Python >=3.7.x
```

## Getting Started
Two files are required: **Sender** and **Recipient**. These can be any file extension.  
- *Currently, only CSV and Excel (.csv, .xlsx) are supported.*  

### Important
Please use the headers shown under the **Data File Structure** section as guidance.

CSV File:
* Will only read and process the second row. 
* The first row, header, will be ignored.  

Excel File:
* Will only read and process from the second row onwards.
* The first row, header, will be ignored.
* The **sheet name** (*also known as the* ***tab name***) must be provided when prompted.

*Full file paths are intended to be used for both inputs and attachment referencing.*

---

### Data File Structure

*Sample **Sender** and **Recipient** header files are provided under the **dataSample** folder.*

##### Sender Data File

Service | Port | Your Login Email | Your Password | Signature
------- | ---- | ---------------- | ------------- | ---------
Required | Required | Required | Optional <li>Depends on the SMTP server</li><li>If omitted, the program will attempt to log in without your password/App password.</li>| Optional  <li>***Must*** be provided with a full path to an .html file</li>

Note that you can set your password with the following if you do not want to use the file:
```python
>>> email_writer.EmailWriter().set_pass('Your Password')
```

##### Recipient Data File  

To addresses | CC addresses | Subject | Body | Attachments
------------ | ------------ | ------- | ---- | -----------
Required <li>Separated by commas</li> | Optional <li>Separated by commas</li> | Optional | Optional <li>Use "\n" for new lines</li> <li>Must be plain text</li> <li> See parameters.py for in-text customizations such as highlighting</li> | Optional <li>Separated by commas</li> <li>Full file path must be provided.</li>

---

### Port Configurations
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

---

### Reference
*Only for the Sender file*
>**Port** - *587*

*Services*
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

### Procedure

1. Launch program from ***run.py***.
2. Import the requested ***Recipient*** data file path with elements corresponding to the **Data File Structure**
3. Import the requested ***Sender*** data file path with elements corresponding to the **Data File Structure**
4. Emails sent will be displayed on the terminal.

---

## Credits
*Code citations are available in-line.*

## Author
Steven Au

## License
See the LICENSE.md for details