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

def the_email_details():
    """Email detail inporter"""
    while 1:
        each_email_detail = mr.file_inputer_csv()
        each_email_detail = each_email_detail[0] # Quick correction

        service = each_email_detail[0]
        port = int(each_email_detail[1])
        sender = each_email_detail[2]
        your_pass = each_email_detail[3]

        try:
            server = smtplib.SMTP(service, port)
        except:
            print("Could not connect to your requested Service and Port", service, port)
            print('Restarting from the beginning - please save a new CSV file and re-enter the file path accordingly.')
            continue

        # IF Login Auth is needed
        if your_pass != 'nothing here': # login
            server.ehlo
            server.starttls()
            try:
                server.login(sender,your_pass)
            except:
                print('Failed to log in with your email and password.')
                print('Restarting from the beginning - your creditionals were not accepted. Please ensure that you are using the App Password Feature.')
                continue             
        return sender, server

def email_fill_in(from_user, server, each_element): # Add optional signature parameter
    """
    Send emails to the appropriate recipients
    Function relies on the following format with headers:
    "To email(s) | CC email(s) | Email subject | Email body | Email attachment(s) | Email signature"
    """
    msg = MIMEMultipart()
    sender = str(from_user)
    to_email = str(each_element[0])
    cc_email = str(each_element[1])
    e_subject = str(each_element[2])
    body = str(each_element[3])
    filename = str(each_element[4])
    signa = str(each_element[5]) # optional
    
    try:
        # Main Address Fields
        msg['From'] = sender
        if to_email.find("@") == -1:
            1/0 # Quickly raise the exception
        msg['To'] = to_email
        if e_subject == 'nothing here':
            e_subject = '[No Subject]'
        msg['Subject'] = e_subject
        if cc_email != "nothing here" or cc_email.find('@') != -1:
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
                continue # If a file fails, go to the next one

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
    """Get attachment type and return it for attachment purposes"""
    
    file_name = filename.split('\\')[-1:]   # Easy name read for attachment purposes at the end
    filename = filename.strip()
    """
    Credit below in reference to the following StackOverflow response.
    https://stackoverflow.com/questions/23171140/how-do-i-send-an-email-with-a-csv-attachment-using-python
    """
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
    """Execute the program per the mailing list data inputs."""
    continue_marker = 'y'
    while continue_marker == 'y':
        sender, server = the_email_details()
        print('Email Number: Email To | CC (Can be blank) | Subject | Body | Attachment path with Extension | optional: Signature path in .html')
        for each_element in range(len(mailinglist)):
            email_fill_in(sender, server, mailinglist[each_element])
            print('The following email was sent:')
            print('Email #'+str(each_element+1)+': '+'|'.join(map(str,mailinglist[each_element]))) # Utilize the join and map interable for list -> str per concat
        print('Do you want to continue sending another email batch?')
        continue_marker = input('If so, please input "y" (any case, no space), otherwise any other key will end the email automation.').lower()
        if continue_marker == 'y':
            continue
        else:
            break

if __name__ == '__main__':
    mailinglist = mr.start_prompt_with_sheets() # This can execute as its own program itself, given the modules
    auto_email_sender(mailinglist)

