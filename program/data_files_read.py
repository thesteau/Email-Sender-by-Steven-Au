# Reads the data file for use with the mailing.
import os

import pandas as pd
from program import parameters as params

class ReadMailing:

    def __init__(self):
        self._target_value = params.Parameters().send_target()
        self._the_word = params.Parameters().send_word()
        self._valid_extensions = params.Parameters().send_extensions()

    def file_read(self):
        read_file_path = input("> ")

        if not os.path.isfile(read_file_path):
            print('No valid file entered. Please try again.')
            return self.file_read()

        return read_file_path

