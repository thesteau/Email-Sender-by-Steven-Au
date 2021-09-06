# Reads the data file for use with the mailing.
import os

import pandas as pd


class ReadMailing:

    def __init__(self):
        self._target_value = 5

    def recitals_recipients(self):
        """ Instructions displayed to the user per the recipients file."""
        print("Important, you must have the following header convention (names can be anything) "
              "with data to use this program:")
        print('Email To | CC (Can be blank) | Subject | Body | Attachment path with Extension')
        print('Use a comma "," to separate more than one To, CC, or Attachments')
        print('You could drag and drop the file if this was launched via the Python terminal')
        print("Alternatively, copy and paste the direct file path with file name. "
              "\nExample: C:Users\\Dummy sheet.xlsx\nThen paste it here.")

    def recitals_sender(self):
        """ Instructions displayed to the user per the sender file."""
        print("For auto to work, please provide the file link for the following in CSV format on the 2nd row: ")
        print("SMTP server | SMTP Port | Your_Email | Your_Password "
              "(Leave blank if not needed) | optional: Signature path in .html")
        print('Please ask your IT department if needed. Type "break" to default\n')

