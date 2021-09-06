# Author: Steven Au
# Purpose: Easily use elements from a CSV and an Excel files to automatically send bulk emails.
# import files here

def main():
    """Emailing choice to automatically or GUI manual process emails."""
    print('Welcome to the automatic email program by Steven Au!')
    print('Please have the following files prepared: \n1.Your emailing credentials\n2. The recipient file.')
    print('Both files must have the details filled in accordingly per the headers.')
    input('Press enter to proceed...')

    active_runner = 'y'
    while active_runner == 'y':
        # email program run

        print()
        print('Do you want to send further emails?')  # Assumes both the credentials and recipient files will differ.
        active_runner = input("Input [y]: yes to send further emails, else any other key for no.\n>").lower()

        if active_runner == 'y':
            print('Restarting now....')
        else:
            print('Aborting now.....')


if __name__ == "__main__":
    main()
