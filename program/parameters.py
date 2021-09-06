# Data file of values for Composition.

class Parameters:
    """ Additional sub data for processing."""

    # Reference plaintext colors and fonts.
    PURPLE = '\033[95m'
    CYAN = '\033[96m'
    DARKCYAN = '\033[36m'
    BLUE = '\033[94m'
    GREEN = '\033[92m'
    YELLOW = '\033[93m'
    RED = '\033[91m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'
    END = '\033[0m'

    def __init__(self):
        self._target_value = 5
        self._the_word = '~&^~nothing here~$%~&'
        self._valid_extensions = [
            '.csv',
            '.xlsx'
        ]

    def get_target(self):
        """ Send the target column value."""
        return self._target_value

    def get_word(self):
        """ Send the key default word (Different from Python's special None keyword."""
        return self._the_word

    def get_extensions(self):
        """ Send the valid extension value list."""
        return self._valid_extensions

    def recitals_program_start(self):
        """ Program welcome to user."""
        print('Welcome to the automatic email program by Steven Au!\n')
        print('Please have the following files prepared: \n1. Your emailing credentials\n2. The recipient file.')
        print('Both files must have the details filled in accordingly per the headers.\n')
        print('Please see the ReadMe for details.')
        input('Press any key to proceed...')
        print()

    def recitals_recipients(self):
        """ Instructions displayed to the user per the recipients file."""
        print("Loading recipient data")
        print("Important, you must have the following header convention (names can be anything) "
              "with the corresponding data to use this program:")
        print('Email To | CC (Can be blank) | Subject | Body | Attachment path with Extension')
        print('Use a comma "," to separate more than one To, CC, or Attachments')
        print("Alternatively, copy and paste the direct file path with file name. "
              "\nExample: C:Users\\Dummy sheet.xlsx\nThen paste it here.")

    def recitals_sender(self):
        """ Instructions displayed to the user per the sender file."""
        print("Loading the sender data.")
        print("Please provide the file path for the following on the 2nd row: ")
        print("SMTP server | SMTP Port | Your_Email | Your_Password "
              "(Leave blank if not needed) | optional: Signature path in .html format.")
        print('Please ask your IT department if needed.')
