# Reads the data file and converts the file to a pandas dataframe for use with the mailing.
import os

import pandas as pd
from program import parameters as params


class ReadMailing:

    def __init__(self):
        self._target_value = params.Parameters().send_target()
        self._the_word = params.Parameters().send_word()
        self._valid_extensions = params.Parameters().send_extensions()

    def file_read(self):
        """ Reads to check if the path entered leads to a file. Else, recursively prompt user to enter again.
            Returns:
                read_file_path = the valid full file path entered by the user.
        """
        read_file_path = input("> ")

        if not os.path.isfile(read_file_path):
            print('No valid file path entered. Please try again.')
            return self.file_read()

        return self.valid_file_check(read_file_path)  # Check the file path is valid

    def valid_file_check(self, file_path):
        """ Check if the file path entered valid per the program."""
        # Ends the check if the file type is valid.
        for each_check in self._valid_extensions:
            if file_path.endswith(each_check):
                return file_path  # Send the file path that was being evaluated.

        # Invalid file type, notify user
        print('This is not a valid file extension. The valid extensions are as followed:')
        for each_ext in self._valid_extensions:
            print(each_ext, end=' ')
        print()
        print('Please try again.')

        return self.file_read()  # Prompt user to reenter

    def standalone(self):
        """ Runs the file with the essential methods as a standalone script."""


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
        print('Please ask your IT department if needed.\n')


test = ReadMailing()
test.file_read()
