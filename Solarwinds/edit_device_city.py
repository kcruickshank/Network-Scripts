import re
import requests
from orionsdk import SwisClient
import openpyxl
import os
import tkinter as tk
from tkinter import filedialog
import warnings
import getpass


def main():
    npm_server = '10.251.6.63'
    #username = 'kenneth.cruickshank'
    username = input('Enter Normal AD Username: ')
    password = getpass.getpass('Enter Password: ')
    verify = False
    if not verify:
        from requests.packages.urllib3.exceptions import InsecureRequestWarning
        requests.packages.urllib3.disable_warnings(InsecureRequestWarning)
    swis_npm = SwisClient(npm_server,username,password)
    print("Getting file update details..")
    file = get_file()
    custom_update = read_custom(file)  
    update_custom(swis_npm,custom_update)
    print("\nCompleted")

def get_file():
    # Create Dialogue file to select xlsx file
    #user_dir = os.getenv("USERPROFILE")
    user_dir = 'C:\TFTP-Root\Solarwinds'
    root = tk.Tk()  # Creates root window
    root.withdraw()  # Hides the root window so only the file dialog shows
    file_options = {'initialdir':user_dir, 'filetypes':[('Excel Files','.xlsx'),('All FIles','.*')]}
    file_path = filedialog.askopenfilename(**file_options)  
    if not file_path:
        raise SystemExit  # If no file chosen exit the program
    return file_path

def read_custom(file):
    warnings.simplefilter('ignore')
    wb = openpyxl.load_workbook(file)
    sheet = wb['Edit City']
    custom_updates = [{
        'Caption':sheet['A'+str(row)].value,
        'City':sheet['B'+str(row)].value
        } for row in range(2,sheet.max_row + 1)]

    print('\nFinished getting updates for City')
    #input('\nPress Enter to continue')
    return custom_updates
    
def update_custom(swis_npm,props2):
    print("Starting to update City details..")
    for item in props2:
        caption = '{}'.format(item['Caption'])
        device_city = '{}'.format(item['City'])
        
        # Check to make sure there is an old and a new value
        if caption == 'None':
            print('Skipping edit as node name is missing')
            continue
        elif device_city == 'None':
            print('Skipping edit as city is blank')
            continue
        else:
            try:
                results = swis_npm.query(
                "SELECT Uri FROM Orion.Nodes WHERE Caption = @caption",
                caption=caption)  # Get Uri!

                print('Updating Device City to ' + device_city + ' for Node ' + caption)
                uri = results['results'][0]['Uri']
                swis_npm.update(uri + '/CustomProperties', City=device_city)
                continue
            except IndexError:
                print('Node {} to be updated does not exist'.format(item['Caption']))
                continue

requests.packages.urllib3.disable_warnings()

    
if __name__ == '__main__':

    main()