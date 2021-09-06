# Data file of values for Composition.

class Parameters:

    def __init__(self):
        self._target_value = 5
        self._the_word = 'nothing here'
        self._valid_extensions = [
            '.csv',
            '.xlsx'
        ]

    def send_target(self):
        return self._target_value

    def send_word(self):
        return self._the_word

    def send_extensions(self):
        return self._valid_extensions


