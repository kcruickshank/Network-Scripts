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
    print("Getting file details..")
    file = get_file()
    file_details = read_file(file)
    write_sheet(file_details)
    print("\nPaste file completed")

def get_file():
    # Create Dialogue file to select xlsx file
    #user_dir = os.getenv("USERPROFILE")
    user_dir = 'C:\TFTP-Root\Wireless\Final Output'
    root = tk.Tk()  # Creates root window
    root.withdraw()  # Hides the root window so only the file dialog shows
    file_options = {'initialdir':user_dir, 'filetypes':[('Excel Files','.xlsx'),('All FIles','.*')]}
    file_path = filedialog.askopenfilename(**file_options)  
    if not file_path:
        raise SystemExit  # If no file chosen exit the program
    return file_path

def read_file(file):
    warnings.simplefilter('ignore')
    wb = openpyxl.load_workbook(file)
    sheet_name = input('Enter the Name of the sheet: ')
    sheet = wb[sheet_name]
    lst_updates = [{
        'AP':sheet['A'+str(row)].value,
        } for row in range(1,sheet.max_row + 1)]

    print('\nFinished getting AP Names\n')
    #input('\nPress Enter to continue')
    return lst_updates
    
def write_sheet(props):
    ip_tftp_server = "10.251.6.35"
    tftp_path_final = "\\" + "\\" + str(ip_tftp_server) + '\TFTP-Root\\Wireless\\Final Output\\'
    site_name = input('Enter Name of Site: ')
    prim_wlc_ip = input('Enter Primary WLC IP: ')
    prim_wlc_name = input('Enter Primary WLC Name: ')
    sec_wlc_ip = input('Enter Secondary WLC IP: ')
    sec_wlc_name = input('Enter Secondary WLC Name: ')
    print("Creating new paste file..")
    # Create a new temp text files
    txt_file_1=open(tftp_path_final + str(site_name) + ' WLC Failover Pastes.txt', "w")
    # Write the heading to text file and close the file
    txt_file_1.write('Paste for Failing APs to the Primary WLC...\n\n')
        
    for item in props:
        i = 1
        while i <= 3:
            if i == 1:
                txt_file_1.write('config ap secondary-base TEMP {} 10.10.10.10\n'.format(item['AP']))
                i += 1
            elif i == 2:
                txt_file_1.write('config ap primary-base ' + str(prim_wlc_name) + ' {} '.format(item['AP']) + str(prim_wlc_ip) + '\n')
                i += 1
            elif i == 3:
                txt_file_1.write('config ap secondary-base ' + str(sec_wlc_name) + ' {} '.format(item['AP']) + str(sec_wlc_ip) + '\n')
                i += 1
    
    # Write the heading to text file 
    txt_file_1.write('\nPaste for Failing APs to the Secondary WLC...\n\n')

    for item in props:
        i = 1
        while i <= 3:
            if i == 1:
                txt_file_1.write('config ap secondary-base TEMP {} 10.10.10.10\n'.format(item['AP']))
                i += 1
            elif i == 2:
                txt_file_1.write('config ap primary-base ' + str(sec_wlc_name) + ' {} '.format(item['AP']) + str(sec_wlc_ip) + '\n')
                i += 1
            elif i == 3:
                txt_file_1.write('config ap secondary-base ' + str(prim_wlc_name) + ' {} '.format(item['AP']) + str(prim_wlc_ip) + '\n')
                i += 1

    txt_file_1.close()

if __name__ == '__main__':

    main()