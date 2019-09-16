"""Program written by Steven Au"""
import pyautogui as pya
import os
import time
import pya_key_combos as pkc

## Manual Method    
def wind_mail_log():
    """Open Gmail"""
    pya.PAUSE = 1.5
    pya.press('win')
    pya.typewrite('Google Chrome')
    pya.press('enter')
    pya.PAUSE = 5
    pya.typewrite('https://mail.google.com/')
    pya.press('enter')

### Gmail Page
def new_gmail_page():
    """New email"""
    pya.PAUSE=5
    pya.press('f5')
    pya.typewrite('d')
    pya.PAUSE=3

### Email construction
def emailing_loop_method(mailinglist):
    """Emails"""
    
    def attachment_procedure(element):
        """Email Attachment"""
        pya.typewrite(element, interval = 0.1) 
        pya.PAUSE = 3
        pya.press('enter')
        pya.PAUSE = 5

    def cc_emails(element):
        """CC emails"""
        pkc.control_shift_combination("c")
        pya.typewrite(element, interval = 0.1)
    
    pya.PAUSE = 5
    for eachrecord in mailinglist:
        new_gmail_page()
        
        ## subject
        pya.PAUSE = 1.5
        pya.press('tab')
        pya.typewrite(eachrecord[2])

        ## body
        pya.PAUSE = 1.5
        pya.press('tab')
        pya.typewrite(eachrecord[3])

        ## file attachment block
        if eachrecord[4] == 'nothing here':
            pass
        else:
            pya.PAUSE = 1.5
            pya.press('tab')
            pya.press('tab')
            pya.press('tab')
            pya.press('enter')
            attachment_procedure(eachrecord[4])

        ## Recipient
        """body to recipient field"""
        pya.PAUSE = 1.5
        pya.keyDown('shift')
        pya.press('tab')
        pya.press('tab')
        pya.keyUp('shift')
        pya.typewrite(eachrecord[0].replace(' ','                 '), interval = 0.1) 

        ## cc block ; no bcc
        if eachrecord[1] == 'nothing here':
            pass
        else:
            cc_emails(eachrecord[1].replace(' ','                 ')) 

        ## close email
        pya.PAUSE = 1.5
        pya.press('tab')
        pkc.control_key_combination('w')
        pya.press('enter')

def manual_gmail_method(mailinglist):
    wind_mail_log()
    emailing_loop_method(mailinglist)

if __name__ == "__main__":
    import Mailing_list_read as mr
    mailinglist = mr.start_prompt_with_sheets()
    manual_gmail_method(mailinglist)
    
