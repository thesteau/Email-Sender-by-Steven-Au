# Program written by Steven Au

import Manual_gmail_draft as md
import Mailing_list_read as mr
import Emailing_Automation_Program as eap

def main():
    """Emailing choice to automatically or GUI manual process emails."""
    print('Welcome to the email program!')
    while True:
        mailinglist = mr.start_prompt_with_sheets() # Grab Data
        choice_check = input("a = auto, any for manual: ")

        if choice_check == "a":
            eap.auto_email_sender(mailinglist) # Send based on parameters
        else: 
            md.manual_gmail_method(mailinglist) # Get in Gmail

        print('')
        print('Do you want to return to the main menu to selectt either Auto or Manual?')
        againer = input("Input (y)es or a(n)y other key for no.\n>").lower()
        if againer == "y":
            print('Restarting now....')
            continue
        else:
            print('Aborting now.....')
            break
            
### The program
if __name__ == "__main__":
    main()
