# Bulk Email Sender by Steven-Au
Automatically send bulk emails based on data elements for both the Sender and Recipient(s) using data files.
  
The Recipient file can be customized and scaled to the number to email receivers and file attachments with a personalized/independent body message each due to the data element stored in the Recipient data file. Please see the **Data File Structure** section for more details regarding how data is to be imported.

## Requirements
```
pandas 1.3.2
```

## Getting Started
Two files are required: **Sender** and **Recipient**.  
- These two files can be any file extension.  
  - *Currently, the following file types are supported as a Data File:
    - .csv (CSV)
    - .xlsx (Excel)

### Important
Please use the headers shown under the **Data File Structure** section as guidance to your data files.

CSV File:
* Will read and process from the second row onwards. 
* The first row, header, will be ignored.  

Excel File:
* Will read and process from the second row onwards. 
* The first row, header, will be ignored.
* The **sheet name** (*also known as the* ***tab name***) must be provided when prompted.

*Full file paths are intended to be used for both inputs and attachment referencing.*

---

### Data File Structure

*Sample **Sender** and **Recipient** header files are provided under the **dataSample** folder.*

##### Sender Data File

Service | Port | Your Login Email | Your Password | Signature
------- | ---- | ---------------- | ------------- | ---------
Required | Required | Required <li>Must contain ```@```symbol</li> | Optional <li>Depends on the SMTP server</li><li>If omitted, the program will attempt to log in without your password/App password.</li>| Optional  <li>A full file path to the ```.html``` file ***must*** be provided</li>

**Note**: You can set your password via the following if you want to access the ```EmailWriter()``` independently:
```python
>>> from program import email_writer
>>> email_writer.EmailWriter().set_pass('Your Password')
```

##### Recipient Data File  

To Addresses | CC Addresses | Subject | Body | Attachments
------------ | ------------ | ------- | ---- | -----------
Required <li>Separated by commas</li><li>Must contain ```@``` symbol</li>  | Optional <li>Separated by commas</li><li>Must contain ```@``` symbol</li>  | Optional | Optional <li>Use "\n" for new lines</li> <li>Must be plain text</li> <li> See parameters.py for in-text customizations such as highlighting</li> | Optional <li>Separated by commas</li> <li>Full file path must be provided.</li>

**Note**: The file attachments can be a mix of any file types such as:
```
.mp3, .mp4, .xlsx, .xls, .csv, .py, .html, .css, .js, .txt, .pptx, .docx, etc.
```

---

### Port Configurations
The encryption used in this program is TLS and ***not*** SSL 

So the Python code uses the following:
```python
>>> server = smtplib.SMTP(service, port)
```
Otherwise, SSL is simply the following for reference:
```python
>>> server = smtplib.SMTP_SSL(service, port)
```

Where ```service``` and ```port``` are your smtp service and port used.

---

### Reference
*Only for the **Sender** file*
>**Port** | *587* - Note that this is an ```integer```

*Services* - These are ```string``` elements.
>**Outlook** | 
*smtp-mail.outlook.com*

>**Yahoo** |
*smtp.mail.yahoo.com*

>**Google** |
*smtp.gmail.com*

>**Hotmail** |
*smtp.live.com*

https://support.google.com/a/answer/176600?hl=en

---

### Procedure

1. Launch the program via ***run.py***.
2. Import the requested ***Recipient*** data file path with elements corresponding to the header from **Data File Structure**
3. Import the requested ***Sender*** data file path with elements corresponding to the header from **Data File Structure**
4. Once confirmed, emails that are sent will be displayed on the terminal.

---

## Program Structure
```graphql
run.py                       # Main program execution
  └─ ./program
     ├─ __init__.py
     ├─ data_files_read.py   # Reads and exports data files into a python list via pandas.
     ├─ email_writer.py      # Log into service, Write(s) and send(s) emails based on list elements.
     └─ parameters.py        # Program reference data and functions through composition.
```

## Credits
*Code citations are available in-line.*

## Author
Steven Au

## License
See the LICENSE.md for details
