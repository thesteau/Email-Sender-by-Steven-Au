# Author: Steven Au
# Purpose: Easily use elements from a CSV and an Excel files to automatically send bulk emails.
# import files here
from program import parameters as params
from program import data_files_read as dfr
from program import email_writer as ew

def main():
    """ Email program primary runner."""
    p = params.Parameters()
    recipients = dfr.ReadMailing()
    sender = dfr.ReadMailing()

    p.recitals_program_start()

    active_runner = 'y'
    while active_runner == 'y':
        # email program run

        p.recitals_recipients()
        recipients.data_read()
        print('----')
        p.recitals_sender()
        sender.data_read()

        print(recipients.show_pandas())
        print('setgrdgrtfghg')
        print(sender.show_pandas())

        print()
        print('Do you want to send further emails?')  # Assumes both the credentials and recipient files will differ.
        active_runner = input("Input [y]: yes to send further emails, else any other key for no.\n> ").lower()

        if active_runner == 'y':
            print('Restarting now....')
        else:
            print('Aborting now.....')


if __name__ == "__main__":
    main()
