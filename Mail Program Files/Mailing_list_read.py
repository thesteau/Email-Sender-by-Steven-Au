import pandas as pd
import xlrd
import os

def start_prompt_with_sheets():
    """Grab sheet and store into a Pandas Dataframe"""
    
    # Recitals
    print("Important, you must have the following header convention (names can be anything) with data to use this program:")
    print('Email To | CC (Can be blank) | Subject | Body | Attachment path with Extension | optional: Signature path in .html')
    print('Use a comma "," to seperate more than one To, CC, or Attachments')
    print('You could drag and drop the file if this was launched via the Python terminal')
    print("Alternatively, copy and paste the direct file path with file name. \nExample: 'C:Users\\Dummy sheet.xlsx'\nThen paste it here:")

    xlfile = input("> ")
    sheetname = input("Now type or paste the exact sheet name. This is case sensitive!\n> ")
    while True:
        try:
            xl = pd.ExcelFile(xlfile)
            df = pd.read_excel(xl, sheetname).fillna('nothing here')
            print('')
            return df.values.tolist()  # all values are now exported as a list 
        except:
            print("Error: Enter the path and sheet name again!")
            xlfile = input("Excel Path (You can drag and drop for the path to be copied)\n> ")
            sheetname = input("Sheet name; type or paste. This is case sensitive!\n> ")
            continue

def file_inputer_csv():
    """Grab sheet and read as a CSV via Pandas"""
    
    # Recitals
    print("For auto to work, please provide the file link for the following in CSV format on the 2nd row: ")
    print("SMTP server | SMTP Port | Your_Email | Your_Password (Leave blank if not needed)\n")
    print("Please ask your IT department if needed.\n")

    the_csv = input("> ")
    while True:
        try:
            cv = pd.read_csv(the_csv).fillna('nothing here')
            print('')
            return cv.values.tolist()
        except:
            print("\nError reading the file path, try again!")
            print("Please note that this has to be a CSV file with a valid full path\n")
            print("SMTP server | SMTP Port | Your_Email | Your_Password (Leave blank if not needed)")
            the_csv = input("> ")
            continue

def base_func():
    """Base function to execute the procedure"""
    while True:
        format_type = input("C for CSV, E for Excel\n").lower()
        if format_type == "e":
            print(start_prompt_with_sheets())
        elif format_type == "c":
            print(file_inputer_csv())
        else:
            print("no key selected, aborting")
            break

if __name__ == "__main__":
    base_func()
