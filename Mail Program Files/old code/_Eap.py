## Secondary

import smtplib
import time
import mimetypes
import os

import pandas as pd
import xlrd

from email.message import Message
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.audio import MIMEAudio
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase
from email import encoders

class SheetDataImport:
    """"""

    def __init__(self):
        self._target_number_e = 5
        self._target_number_c = 5

    def prompts(self, value=None):
        """"""

        if value is not None:

            # Prompt value matrix
            if value == 1:
                print('Type "break" to default.')
                print("Important, you must have the following header convention (names can be anything) with data to use this program:")
                print('Email To | CC (Can be blank) | Subject | Body | Attachment path with Extension')
                print('Use a comma "," to seperate more than one To, CC, or Attachments')
                print('You could drag and drop the file if this was launched via the Python terminal')
                print("Alternatively, copy and paste the direct file path with file name. \nExample: C:Users\\Dummy sheet.xlsx\nThen paste it here.")

            if value == 2:
                print("For auto to work, please provide the file link for the following in CSV format on the 2nd row: ")
                print(
                    "SMTP server | SMTP Port | Your_Email | Your_Password (Leave blank if not needed) | optional: Signature path in .html")
                print('Please ask your IT department if needed. Type "break" to default\n')

    def excel_import(self):
        # Use the global number for list default - note that the list may vary if a column is either added or removed

        self.prompts(1)

        xlfile = self.empty_func_list(self._target_number_e)
        if xlfile == [['nothing here'] * self._target_number_e]:
            return xlfile
        sheetname = input("Now type or paste the exact sheet name. This is case sensitive!\n> ")
        while 1:
            try:
                xl = pd.ExcelFile(xlfile)
                df = pd.read_excel(xl, sheetname).fillna('nothing here')
                print('')
                return df.values.tolist()  # all values are now exported as a list
            except:
                print("Error: Enter the path and sheet name again!")
                print("Excel Path (You can drag and drop for the path to be copied)\n")
                xlfile = self.empty_func_list(self._target_number_e)
                if xlfile == [['nothing here'] * self._target_number_e]:
                    return xlfile
                sheetname = input("Sheet name; type or paste. This is case sensitive!\n> ")
                continue

    # Comma Seperated Variable
    def file_inputer_csv(self):
        """Grab sheet and read as a CSV via Pandas"""
        # Recitals

        self.prompts(2)

        # Use the global number for list default - note that the list may vary if a column is either added or removed
        the_csv = self.empty_func_list(self._target_number_c)
        if the_csv == [['nothing here'] * self._target_number_c]:
            return the_csv
        while 1:
            try:
                cv = pd.read_csv(the_csv).fillna('nothing here')
                print('')
                return cv.values.tolist()
            except:
                print("\nError reading the file path, try again!")
                print("Please note that this has to be a CSV file with a valid full path\n")
                print(
                    "SMTP server | SMTP Port | Your_Email | Your_Password (Leave blank if not needed) | optional: Signature path in .html\n")
                the_csv = self.empty_func_list(self._target_number_c)
                if the_csv == [['nothing here'] * self._target_number_c]:
                    return the_csv
                continue

    def empty_func_list(self, target_number):
        """Use the global number for list default - note that the list may vary if a column is either added or removed"""
        what_is_said = input("> ")
        what_is_said.replace('"', '')
        if what_is_said.lower() == 'break':
            emptiness = [['nothing here'] * target_number]
            return emptiness
        else:
            return what_is_said
## Finished


class EmailAuto:
    """"""

    def __init__(self, mailing_list):
        """"""
        self._the_word = "nothing here"
        self._mailing_list = mailing_list

    def _element_check(self, list_check, number):
        """"""
        try:
            the_element = str(list_check[number])
        except:
            the_element = self._the_word
        return the_element

    def the_email_details(self):
        """"""

        while True:
            each_email_detail = None
            break



