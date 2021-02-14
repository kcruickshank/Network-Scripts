import re
import openpyxl
import os
import tkinter as tk
from tkinter import filedialog
import warnings
import getpass

new1= "kc"
new12= "kc2"

def main():
    file = get_file()
    file_contents = read_file(file)
    print(file_contents)
    #update_custom(file_contents)
    print('\nCompleted.')

def get_file():
    # Create Dialogue file to select xlsx file
    #user_dir = os.getenv("USERPROFILE")
    start_dir = 'C:\Test_Files'
    root = tk.Tk()  # Creates root window
    root.withdraw()  # Hides the root window so only the file dialog shows
    file_options = {'initialdir':start_dir, 'filetypes':[('Excel Files','.xlsx'),('All FIles','.*')]}
    file_path = filedialog.askopenfilename(**file_options)  
    if not file_path:
        raise SystemExit  # If no file chosen exit the program
    return file_path

def read_file(file):
    warnings.simplefilter('ignore')
    wb = openpyxl.load_workbook(file)
    sheet = wb['Sheet1']
    lst_updates = [{
        'ip_address':sheet['A'+str(row)].value,
        'switch_name':sheet['B'+str(row)].value
        } for row in range(2,sheet.max_row + 1)]
    return lst_updates

def update_custom(sheet_details):
    print("\nStarting to Update Names..")
    for item in sheet_details:
        print('\nLogging into switch IP Address {}'.format(item['ip_address']))
        print('Changing name to {}'.format(item['switch_name']))
        
    
if __name__ == '__main__':

    main()