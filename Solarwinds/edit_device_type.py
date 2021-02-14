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
    username = 'kenneth.cruickshank'
    #username = input('Enter Username: ')
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
    sheet = wb['Edit Device Type']
    custom_updates = [{
        'Caption':sheet['A'+str(row)].value,
        'Device_Type':sheet['B'+str(row)].value
        } for row in range(2,sheet.max_row + 1)]

    print('\nFinished getting updates for Device Type')
    return custom_updates
    
def update_custom(swis_npm,props2):
    print("Starting to update Device Type details..")
    for item in props2:
        caption = '{}'.format(item['Caption'])
        device_type = '{}'.format(item['Device_Type'])
        
        # Check to make sure there is an old and a new value
        if caption == 'None':
            print('Skipping edit as node name is missing')
            continue
        elif device_type == 'None':
            print('Skipping edit as device type is blank')
            continue
        else:
            try:
                results = swis_npm.query(
                "SELECT Uri FROM Orion.Nodes WHERE Caption = @caption",
                caption=caption)  # Get Uri!

                print('Trying to Update the Device Type to {} for Orion Node {}'.format(item['Device_Type'],item['Caption']))
                uri = results['results'][0]['Uri']
                swis_npm.update(uri + '/CustomProperties', DeviceType=device_type)
                continue
            except IndexError:
                print('\n ERROR - Node {} does not exist in Orion with that name'.format(item['Caption']))
                continue

requests.packages.urllib3.disable_warnings()

    
if __name__ == '__main__':

    main()
        