## Program was written by Steven Au
# Python 3.7.4
# coding: utf-8
import smtplib
import time
import mimetypes
from email.message import Message
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.audio import MIMEAudio
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase
from email import encoders
import Mailing_list_read as mr

"""
DOC

to do:
    Update to docstrings to Google format
    Change the importing of the reader script to only import without much of its verbage or to redirect accordingly to allow script reuse in other projects

    Add a function to simply replace the NaN from Pandas instead of using the .fillna, thus allowing the script to be independent for future projects

    Change the reader module to an alternative name
    Add a None default to the mailing_list_reader module solely for independent recitals if any


    Add a custom delay time via the time module

"""


## Global Constant - Error Word
THE_WORD = 'nothing here'

def element_check(listchk, number):
    """Element checker

    This checks 
    This is to test whter the sheets used contain the requested elements
    
    Args:
        listchk: Elements
        number (int):

    Returns: 

    Example:



    """
    global THE_WORD
    try:
        the_element = str(listchk[number])
    except:
        the_element = THE_WORD
    return the_element

def the_email_details():
    """Email detail inporter

    details

    Returns:
        sender:
        server:
        signature:
    
    Example:

    
    """
    global THE_WORD

    while 1:
        each_email_detail = mr.file_inputer_csv()
        each_email_detail = each_email_detail[0] # Quick list correction

        service = element_check(each_email_detail, 0)
        port = int(element_check(each_email_detail, 1))
        sender = element_check(each_email_detail, 2)
        your_pass = element_check(each_email_detail, 3)
        signature = element_check(each_email_detail, 4)

        try:
            server = smtplib.SMTP(service, port)
        except:
            print("Could not connect to your requested Service and Port", service, port)
            print('Restarting from the beginning - please save a new CSV file and re-enter the file path accordingly.')
            continue

        # IF Login Auth is needed
        if your_pass != THE_WORD: # login
            server.ehlo
            server.starttls()
            try:
                server.login(sender,your_pass)
            except:
                print('Failed to log in with your email and password.')
                print('Restarting from the beginning - your creditionals were not accepted. Please ensure that you are using the App Password Feature.')
                continue             
        return sender, server, signature

def email_fill_in(from_user, server, signature, each_element):
    """Send emails to the appropriate recipients

    Usage:
    Function relies on the following headers in the Excel File:
        "To email(s) | CC email(s) | Email subject | Email body | Email attachment(s)"
    
        The headers will not be imported and are used for referencing purposes.

    To Do:
        Change the process of the reader script to solely import data and export the lists
        Add various formats such as TSV

    Args:
        from_user: 
        server:
        signature:
        each_element: 

    Returns:


    """
    global THE_WORD

    msg = MIMEMultipart()
    sender = from_user
    to_email = element_check(each_element, 0)
    cc_email = element_check(each_element, 1)
    e_subject = element_check(each_element, 2)
    body = element_check(each_element, 3)
    try:
        filename = element_check(each_element, 4)
    except:
        filename = THE_WORD
    signa = signature # optional
    
    try:
        # Main Address Fields
        msg['From'] = sender
        if to_email.find("@") == -1:
            1/0 # Quickly raise the exception
        msg['To'] = to_email
        if e_subject == THE_WORD:
            e_subject = '[No Subject]'
        msg['Subject'] = e_subject
        if cc_email != THE_WORD or cc_email.find('@') != -1:
            msg['Cc'] = cc_email

        ### Body
        msg.attach(MIMEText(body,'plain'))

        ## Signature - Simple do or go
        try:
            html_file = open(signa, 'r', encoding='utf-8')
            signature = html_file.read()
            html_file.close()
            msg.attach(MIMEText(signature,'html'))
        except:
            pass

        ## Attachment
        filename = filename.split(',')
        for each_file in filename:
            try:
                attached_file = email_attach_func(each_file)
                msg.attach(attached_file)
            except:
                if each_file != (THE_WORD or ['']): # The '' is due to the comma giving an empty string, if added accidentally at the end of a path
                    print('The following file failed to attach..')
                    print(each_file)
                continue # If a file fails, proceed onwards. Ignore printing THE_WORD or ''
        # Send
        server.send_message(msg)

    except: # Signature and Body can be too long.
        print('Could not send email with the following information:')
        print('Sender (From):', sender)
        print('To:', to_email)
        print("Cc:", cc_email)
        print('Subject:', e_subject) # If subjects are differentiated from one antoher, will be useful
        print("Attachment(s):", filename)

def email_attach_func(filename):
    """Get attachment type and return it for attachment purposes
    
    Parameters:


    Credits:
        Credit below in reference to the following StackOverflow response.
        https://stackoverflow.com/questions/23171140/how-do-i-send-an-email-with-a-csv-attachment-using-python
    
    """
    
    file_name = filename.split('\\')[-1:]   # Easy name read for attachment purposes at the end
    filename = filename.strip()

    # Credits @ above
    ctype, encoding = mimetypes.guess_type(filename)
    if ctype is None or encoding is not None:
        ctype = "application/octet-stream"

    maintype, subtype = ctype.split("/", 1)
    
    if maintype == "text":
        the_file_type = open(filename)
        attachment = MIMEText(the_file_type.read(), _subtype=subtype)
        the_file_type.close()
    elif maintype == "image":
        the_file_type = open(filename, "rb")
        attachment = MIMEImage(the_file_type.read(), _subtype=subtype)
        the_file_type.close()
    elif maintype == "audio":
        the_file_type = open(filename, "rb")
        attachment = MIMEAudio(the_file_type.read(), _subtype=subtype)
        the_file_type.close()
    else:
        the_file_type = open(filename, "rb")
        attachment = MIMEBase(maintype, subtype)
        attachment.set_payload(the_file_type.read())
        the_file_type.close()
        encoders.encode_base64(attachment)
    
    attachment.add_header('Content-Disposition', 'attachment', filename=file_name[0])
    return attachment

def auto_email_sender(mailinglist):
    """Execute the program per the mailing list data inputs.
    
    Parameters:


    Process:

    
    """
    continue_marker = 'y'
    while continue_marker == 'y':
        sender, server, signature = the_email_details()
        if 'break' in (sender, server, signature):
            continue
        print('Email Number: Email To | CC (Can be blank) | Subject | Body | Attachment path with Extension')
        for each_element in range(len(mailinglist)):
            email_fill_in(sender, server, signature, mailinglist[each_element])
            print('The following email was sent:')
            print('Email #'+str(each_element+1)+': '+'|'.join(map(str,mailinglist[each_element]))) # Utilize the join and map iterable for list -> str per concat
        print('\nDo you want to continue sending another email batch?')
        continue_marker = input('If so, please input "y" (any case, no space), otherwise any other key to end the email automation.\n').lower()
        if continue_marker == 'y':
            continue
        else:
            break

def main():
    """Main run
    """
    while 1:
        print('If you want to do a log in test, please input "log", otherwise press any key or "enter" to run the program.')
        test = input('> ').lower()
        if test == 'log':
            the_email_details() # Email log in test
            print('Successfully logged in, restarting the process.')
            continue # restart from the beginning if the log in was succeded or not. 
        mailinglist = mr.start_prompt_with_sheets() # This can execute as its own program itself, given the modules
        auto_email_sender(mailinglist)
        continuer = input('Continue? Enter "y" to restart, otherwise any key or "enter" to end.').lower()
        if continuer == 'y':
            continue
        input('Aborting, please wait......')

if __name__ == '__main__':
    main()
