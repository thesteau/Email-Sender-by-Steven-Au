# Reads the data file and converts the file to a pandas dataframe for use with the mailing.
import os

import pandas as pd
from program import parameters as params


class ReadMailing:
    """ Reads the file and converts the data into a pandas dataframe."""

    def __init__(self):
        self._target_value = params.Parameters().send_target()
        self._the_word = params.Parameters().send_word()
        self._valid_extensions = params.Parameters().send_extensions()
        self._file_path = None
        self._pan_data = None

    # Set the file path
    def file_read(self):
        """ Reads to check if the path entered leads to a file. Else, recursively prompt user to enter again.
            Returns:
                read_file_path = the valid full file path entered by the user.
        """
        read_file_path = input("> ")

        if not os.path.isfile(read_file_path):
            print('No valid file path entered. Please try again.')
            return self.file_read()

        self._file_path = read_file_path
        return self.valid_file_check(self._file_path)  # Check the file path is valid

    def valid_file_check(self, file_path):
        """ Check if the file path extension entered is valid per the program.
            Returns:
                BRANCH:
                    file_path = the entered valid file path
                    self.file_read() = Reprompt user to enter a new file path.
        """
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

    # Set the dataframe values
    def process_data(self):
        """ Processes the data accordingly per the file extension.
            Note: Excel files require the "tabs" to be specified in pandas.
            Returns:
                self._pan_data = pandas dataframes of the file data entered.
        """
        if self._file_path is None:
            print('No file exists to process the data.')
            return
        elif self._file_path.endswith('.xlsx'):
            # Recall that excel has multiple tabs
            pd_dat = self.excel_to_df()
        elif self._file_path.endswith('.csv'):
            pd_dat = self.csv_to_df()
        # Further extension paths can be added here.

        self._pan_data = pd_dat
        return self._pan_data

    # Excel
    def excel_to_df(self):

        # try except
        return 1

    # CSV (Comma Separated Variable)
    def csv_to_df(self):
        return 1

    # Further file types to dataframes can be added below.



test = ReadMailing()
test.file_read()
