# Author: Steven Au
# Purpose: Easily use elements from a CSV and an Excel files to automatically send bulk emails.
from .program import parameters as params
from .program import data_files_read as dfr
from .program import email_writer as ew


def main():
    """ Email program primary runner."""
    p = params.Parameters()
    recipients = dfr.ReadMailing()
    sender = dfr.ReadMailing()

    # Introduce user to the program
    p.recitals_program_start()

    # Email data entry and sending
    active_runner = 'y'
    while active_runner == 'y':

        # Data entry read
        p.recitals_recipients()
        recipients.data_read()
        print('----')
        p.recitals_sender()
        sender.data_read()  # Note: Cannot assume that the recipient file contains sender data if using an Excel file.
        print('----')

        print()
        restart = input('Press any key to send emails, otherwise, enter "a" to restart.').lower()
        print()
        print('----')
        if restart == 'a':
            continue

        # Email sending
        emailing = ew.EmailWriter(sender, recipients)

        try:
            emailing.email_processing()
        except:
            print('Restarting the program....')
            continue

        # Program closure
        print()
        print('Do you want to send further emails?')  # Assumes both the credentials and recipient files will differ.
        active_runner = input("Input [y]: yes to send further emails, else any other key for no.\n> ").lower()

        if active_runner == 'y':
            print('Restarting now....')
            print()
        else:
            print('Aborting now.....')
            input()


if __name__ == "__main__":
    main()
