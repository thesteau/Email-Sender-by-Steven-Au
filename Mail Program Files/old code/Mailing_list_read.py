import pandas as pd
import xlrd
import os

## Global multiple - for quick modification purposes based on what values are newly accepted in the program
# Subject to Excel
target_number_e = 5
# Subject to Comma Seperated Variable
target_number_c = 5

# Excel
def start_prompt_with_sheets():
    """Grab sheet and store into a Pandas Dataframe"""
    global target_number_e
    # Recitals
    print('Type "break" to default.')
    print("Important, you must have the following header convention (names can be anything) with data to use this program:")
    print('Email To | CC (Can be blank) | Subject | Body | Attachment path with Extension')
    print('Use a comma "," to seperate more than one To, CC, or Attachments')
    print('You could drag and drop the file if this was launched via the Python terminal')
    print("Alternatively, copy and paste the direct file path with file name. \nExample: C:Users\\Dummy sheet.xlsx\nThen paste it here.")

    # Use the global number for list default - note that the list may vary if a column is either added or removed
    xlfile = empty_func_list(target_number_e)
    if xlfile == [['nothing here']*target_number_e]:
        return xlfile
    sheetname = input("Now type or paste the exact sheet name. This is case sensitive!\n> ")
    while 1:
        try:
            xl = pd.ExcelFile(xlfile)
            df = pd.read_excel(xl, sheetname).fillna('nothing here')
            print('')
            return df.values.tolist()  # all values are now exported as a list 
        except:
            print("Error: Enter the path and sheet name again!")
            print("Excel Path (You can drag and drop for the path to be copied)\n")
            xlfile = empty_func_list(target_number_e)
            if xlfile == [['nothing here']*target_number_e]:
                return xlfile
            sheetname = input("Sheet name; type or paste. This is case sensitive!\n> ")
            continue

# Comma Seperated Variable
def file_inputer_csv():
    """Grab sheet and read as a CSV via Pandas"""
    global target_number_c
    
    # Recitals
    print("For auto to work, please provide the file link for the following in CSV format on the 2nd row: ")
    print("SMTP server | SMTP Port | Your_Email | Your_Password (Leave blank if not needed) | optional: Signature path in .html")
    print('Please ask your IT department if needed. Type "break" to default\n')

    # Use the global number for list default - note that the list may vary if a column is either added or removed
    the_csv = empty_func_list(target_number_c)
    if the_csv == [['nothing here']*target_number_c]:
        return the_csv
    while 1:
        try:
            cv = pd.read_csv(the_csv).fillna('nothing here')
            print('')
            return cv.values.tolist()
        except:
            print("\nError reading the file path, try again!")
            print("Please note that this has to be a CSV file with a valid full path\n")
            print("SMTP server | SMTP Port | Your_Email | Your_Password (Leave blank if not needed) | optional: Signature path in .html\n")
            the_csv = empty_func_list(target_number_c)
            if the_csv == [['nothing here']*target_number_c]:
                return the_csv
            continue

def empty_func_list(global_number):
    """Use the global number for list default - note that the list may vary if a column is either added or removed"""
    what_is_said = input("> ")
    what_is_said.replace('"','')
    if what_is_said.lower() == 'break':
        emptiness = [['nothing here']*global_number]
        return emptiness
    else:
        return what_is_said

def base_func():
    """Base function to execute the procedure"""
    while True:
        format_type = input("C for CSV, E for Excel\n").lower()
        if format_type == "e":
            print(start_prompt_with_sheets())
        elif format_type == "c":
            print(file_inputer_csv())
        else:
            input("No key selected, aborting....")
            break

if __name__ == "__main__":
    base_func()
