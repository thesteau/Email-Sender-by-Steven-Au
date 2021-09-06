## Version 3 Refactoring
# By Steven Au
# Purpose: For use in an office environment for basic email distribution based on data sheet values.

import _Eap


def main():
    """ Email program runner."""
    print("Welcome to the the email automation program!")
    running = True

    while running:
        mailing_list = _Eap.SheetDataImport.excel_import()  # Import the data sheet

        _Eap.EmailAuto(mailing_list)  # Run the automate email program

        print()
        print("Do you need to resend a new set of emails?")
        againer = input('Input "y" for yes or any other key for no and end the program.').strip().lower()

        if againer == 'y':
            print("Restarting now....")
        else:
            print("Aborting now.....")
            running = False


if __name__ == "__main__":
    main()