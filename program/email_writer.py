# Writes and send the emails
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

import pandas as pd
from program import parameters as params


class EmailWriter:
    """ Writes and sends the emails based on data values."""

    def __init__(self, sender, recipients):
        self._the_word = params.Parameters().get_word()
        self._sender = sender
        self._recipients = recipients
        self._server = None

    def email_sender_details(self):
        """ Fills in the sender details.
            Returns:
                server =
                sender =
                signature =
        """
        send_dat = self._sender.get_pandas()[0]

        service = send_dat[0]
        port = int(send_dat[1])
        sender = send_dat[2]
        the_pass = send_dat[3]
        signature = send_dat[4]

        # See the ReadMe for details on the service and port.
        try:
            server = smtplib.SMTP(service, port)
        except:
            print("Could not connect to your requested Service: ", service, ' and Port: ', port)
            print('Restarting from the beginning - please save a new file and re-enter the file path accordingly.')
            raise Exception

        # IF Login Auth is needed
        if the_pass != self._the_word:  # login
            server.ehlo
            server.starttls()

            try:
                server.login(sender, the_pass)
            except:
                print('Failed to log in with your email and password.')
                print('Restarting from the beginning - your credentials were not accepted. '
                      'Please ensure if you are using the App Password Feature to set the password here.')
                raise Exception

        self._server = server
        return sender, signature

    def email_element_valid(self, **kwargs):
        """ Validate individual elements based on keyword arguments."""
        if 'email' in kwargs.keys():
            # If the at symbol is not found, raise an exception
            if kwargs['email'].find('@') == -1:
                raise Exception
        else:  # Invalid call, unavailable in the keyword argument dictionary
            raise Exception

    def email_fill_in(self, from_user, signature, recipient_dat, server=None):
        """ """
        # No server
        if server is None and self._server is None:
            print('No server is currently valid...')
            raise Exception
        elif server is None:
            server = self._server

        msg = MIMEMultipart()

        sender = from_user
        signa = signature  # optional
        to_email = recipient_dat[0]
        cc_email = recipient_dat[1]
        email_subj = recipient_dat[2]
        body = recipient_dat[3]
        filename = recipient_dat[4]

        try:
            msg['From'] = sender
            self.email_element_valid(email=sender)

            msg['To'] = to_email
            self.email_element_valid(email=to_email)

            if email_subj == self._the_word:
                email_subj = '[No Subject]'
            msg['Subject'] = email_subj

            if cc_email != self._the_word or cc_email.find('@') != -1:
                msg['Cc'] = cc_email

            msg.attach(MIMEText(body, 'html'))

            try:
                with open(signa, 'r', encoding='utf-8') as html_sig:
                    msg.attach(MIMEText(html_sig.read(), 'html'))
            except:
                pass

            files = filename.split(',')
            for each_file in files:
                try:
                    attached_file = self.email_attachment(each_file)
                    msg.attach(attached_file)
                except:
                    if each_file != (self._the_word or ['']):
                        # The '' is due to the comma giving an empty string, if added accidentally at the end of a path
                        print('The following file failed to attach..')
                        print(each_file)
                    continue  # If a file fails, proceed onwards. Ignore printing the_word or ''
            server.send_message(msg)

        except:  # Signature and Body can be too long.
            print('Could not send email with the following information:')
            print('Sender (From):', sender)
            print('To:', to_email)
            print("Cc:", cc_email)
            print('Subject:', email_subj)
            print("Attachment(s):", filename)

    def email_attachment(self, filename):
        """ Attach to the email document with the corresponding file types."""
        file_name = filename.split('\\')[-1:]  # Easy name read for attachment purposes at the end
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
            with open(filename) as the_file_type:
                attachment = MIMEText(the_file_type.read(), _subtype=subtype)
        elif maintype == "image":
            with open(filename, "rb") as the_file_type
                attachment = MIMEImage(the_file_type.read(), _subtype=subtype)
        elif maintype == "audio":
            with open(filename, "rb") as the_file_type:
                attachment = MIMEAudio(the_file_type.read(), _subtype=subtype)
        else:
            with open(filename, "rb") as the_file_type:
                attachment = MIMEBase(maintype, subtype)
                attachment.set_payload(the_file_type.read())
                encoders.encode_base64(attachment)

        attachment.add_header('Content-Disposition', 'attachment', filename=file_name[0])
        return attachment

    def email_processing(self, mailing_list):
        print('Email Number: Email To | CC (Can be blank) | Subject | Body | Attachment path with Extension')
        # for each_element in range(len(mailing_list)):
        #     self.email_fill_in(sender, server, signature, mailinglist[each_element])
        #     print('The following email was sent:')
        #     print('Email #'+str(each_element+1)+': '+'|'.join(map(str,mailinglist[each_element])))
        #     # Utilize the join and map iterable for list -> str per concat

