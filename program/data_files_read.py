# Reads and converts the data file into a pandas dataframe for use with the mailing.
# Note openpyxl must be installed for Excel processing.
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

    # Show the current data.
    def show_file_path(self):
        """ Show the currently stored file path."""
        return self._file_path

    def show_pandas(self):
        """ Show the pandas dataframe"""
        return self._pan_data

    # READS
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

    # PROCESSING
    # Set the dataframe values
    def process_data(self):
        """ Processes the data accordingly per the file extension.
            Sets the self._pan_data to the loaded pandas data.
            Returns:
                self._pan_data = pandas dataframes of the file data entered.
        """
        if self._file_path is None:  # No file path is currently stored
            print('No file exists to process the data.')
            return

        # Header row will be skipped.
        self._pan_data = self.data_to_df(self._file_path)
        return self._pan_data

    # Data file to pandas dataframe
    def data_to_df(self, file_path=None, skip_rows=0):
        """ Read and load the data file into a pandas dataframe.
            Note: this method was intended to be "flexible" to load any file into a pandas dataframe on demand.
            Returns:
                df.values.tolist() = pandas dataframe values converted to a list
        """
        if file_path is None:
            print('No data file to process into a dataframe.....')
            return
        elif file_path.endswith(".xlsx"):
            df = self.excel_processing(file_path, skip_rows)
        elif file_path.endswith('.csv'):
            df = self.csv_processing(file_path, skip_rows)
        # Please add additional elif branching for other file extensions.
        else:  # File is currently unsupported.
            print('Please add the corresponding method for this file extension.')
            return

        df.fillna(self._the_word)
        return df.values.tolist()

    # Individual branch processing for the data_to_df method.
    def excel_processing(self, file_path, skip_rows):
        """ Processes the excel file into a pandas dataframe.
            Recursively prompts the user to re-enter the correct sheet name (tab name) if it is erroneous.
            Returns:
                the_dat = the loaded pandas dataframe file.
        """
        # Note: Due to a user entry of the case sensitive sheet name, this check ensures correct processing.
        try:
            print('Note: The tab name is case sensitive, please ensure that it is entered correctly.')
            sheet_name = input('Enter the tab name to be read. > ')
            the_dat = pd.read_excel(file_path, sheet_name, skiprows=skip_rows, engine='openpyxl')
            print('Excel file loaded successfully!\n')
            return the_dat
        except:
            print('The tab name is incorrect.')
            print('Please check the file again for the specific tab name.')
            print('Tip: Copy and paste the tab name.')
            print()
            return self.excel_processing(file_path, skip_rows)

    def csv_processing(self, file_path, skip_rows):
        """ Processes a CSV file into a pandas dataframe.
            Returns:
                the csv pandas dataframe
        """
        print('CSV file loaded successfully!')
        return pd.read_csv(file_path, skiprows=skip_rows)

    # Add additional filetype processing here.
    def tsv_processing(self, file_path, skip_rows):
        """ Processes a TSV file into a pandas dataframe.
            Returns:
                the tsv pandas dataframe
        """
        print('TSV file loaded successfully!')
        return pd.read_csv(file_path, sep='\t', skiprows=skip_rows)

    def data_read(self):
        """ Processes the data with the vital methods as a standalone."""
        self.file_read()
        self.process_data()


if __name__ == "__main__":
    rm = ReadMailing()
    rm.data_read()
    rm.show_pandas()
