import re
import requests
from orionsdk import SwisClient
import openpyxl
import os
from os import path
import tkinter as tk
from tkinter import filedialog
import warnings
import getpass


def main():
    authenticated = False
    counter = 0
    npm_server = '10.251.6.63'

    # Set Directory Path
    dir_path = r"C:\Solarwinds"
    # Check if the Orion Output Spreadsheet exists
    file_path = r"C:\Solarwinds\Orion_Output.xlsx"

    if not path.isfile(file_path):
        print('\nOrion_Output.xlsx does not exist in ' + str(dir_path))
        print('\nLocate the file and save to ' + str(dir_path) + ' or run Get_Nodes.exe again to create New sheet')
        input('\nPress Enter key to Exit')
        raise SystemExit

    print('\n** Use your normal Active Directory Account to connect to Solarwinds **\n')
    
    verify = False
    if not verify:
        from requests.packages.urllib3.exceptions import InsecureRequestWarning
        requests.packages.urllib3.disable_warnings(InsecureRequestWarning)
    
    while not authenticated:
        #username = input('Enter Username: ')
        #password = getpass.getpass('Enter Password: ')
        username = 'mapsacc'
        password = 'nMUDIXKZ3V'
        swis_npm = SwisClient(npm_server,username,password)
        try:
            swis_npm.query("SELECT NodeID FROM Orion.Nodes")
            authenticated = True
        except:
            if counter == 0:
                print("Authentication Error, 1st attempt\n")
                counter += 1
            elif counter == 1:
                print("Authentication Error, 2nd attempt\n")
                counter += 1
            elif counter == 2:
                print("Authentication Error, last attempt\n")
                counter += 1
            else:
                print("Authentication Error, please use a correct username and password\n")
                input("Press any key to exit")
                raise SystemExit
            authenticated = False
    
    ip_address = "10.64.254.20"
    results_dns = swis_npm.query("SELECT ReverseDNS FROM Cirrus.IpAddresses WHERE IPAddress = " + "'" + ip_address + "'")
    #results_dns = swis_npm.query("SELECT IpNodeID, DisplayName  FROM IPAM.IPNode WHERE IPAddress = " + "'" + ip_address + "'")
    # results_dns = swis_npm.query("SELECT Name, DisplayName, Description FROM IPAM.DnsRecordReport")
    #ip_node_id = results_dns['results'][0]['IpNodeID']
    #print(ip_node_id)
    #results1 = swis_npm.query("SELECT DisplayName FROM IPAM.IPNodeGrid WHERE IpNodeId = " + str(ip_node_id) + "")
    # uri_dns = 'swis://localhost/Orion/IPAM.IPNode/IpNodeId=' + str(ip_node_id) + ''
  
    print(results_dns)

    #print(results)
    #print(results1)

    # orion_res = swis_npm.query("SELECT NodeID, FROM Orion.Nodes WHERE IPAddress = " + "'" + ip_address + "'")
    # print(orion_res)

    # ip_add = "10.64.254.25"
    # new_name = "ek03asw001"
    # dns_zone = "cns.muellergroup.com"
    # dns_server_1 = "10.96.65.100"
    # dns_server_2 = '10.96.65.101'
    # dns_server_3 = '10.80.50.45'

    # Add Hostname field
    #swis_npm.update(uri_dns, DnsBackward=new_name + "." + dns_zone)

    ## Add DNS A Record
    # swis_npm.invoke('IPAM.IPAddressManagement', 'AddDnsARecord',new_name,ip_add,dns_server_1, 'cns.muellergroup.com')
    # swis_npm.invoke('IPAM.IPAddressManagement', 'AddDnsARecord',new_name,ip_add,dns_server_2, 'cns.muellergroup.com')
    # swis_npm.invoke('IPAM.IPAddressManagement', 'AddDnsARecord',new_name,ip_add,dns_server_3, 'cns.muellergroup.com')

    ## Remove DNS A Record
    # swis_npm.invoke('IPAM.IPAddressManagement', 'RemoveDnsARecord',new_name + "." + dns_zone + ".",ip_add,dns_server_1,dns_zone)
    # swis_npm.invoke('IPAM.IPAddressManagement', 'RemoveDnsARecord',new_name + "." + dns_zone + ".",ip_add,dns_server_2,dns_zone)
    # swis_npm.invoke('IPAM.IPAddressManagement', 'RemoveDnsARecord',new_name + "." + dns_zone + ".",ip_add,dns_server_3,dns_zone)
     

    #print(results)
    #print(uri)

    #orion_res = swis_npm.query("SELECT Name, IPAddressN FROM IPAM.DnsRecord WHERE Uri = @res",
    #                 res=results)

    #print(orion_res)


    #orion_res = swis_npm.query("SELECT DnsRecordId, DnsZoneId, Name  FROM IPAM.DnsRecord")
    #print(orion_res)

    #file = file_path
    #update_node = read_file(file)
    #update_custom(swis_npm,update_node)

    # Print that the program has completed
    print('\n** Node Names Completed **\n')
    #input('Press enter key to Exit')

# def get_file():
#     # Create Dialogue file to select xlsx file
#     #user_dir = os.getenv("USERPROFILE")
#     #start_dir = 'C:\TFTP-Root\Solarwinds'
#     # Set the start directory to the users desktop
#     start_dir = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
#     root = tk.Tk()  # Creates root window
#     root.withdraw()  # Hides the root window so only the file dialog shows
#     file_options = {'initialdir':start_dir, 'filetypes':[('Excel Files','.xlsx'),('All FIles','.*')]}
#     file_path = filedialog.askopenfilename(**file_options)  
#     if not file_path:
#         raise SystemExit  # If no file chosen exit the program
#     return file_path

def read_file(file):
    warnings.simplefilter('ignore')
    wb = openpyxl.load_workbook(file)
    sheet = wb['Edit Node Name']
    lst_updates = [{
        'Caption':sheet['A'+str(row)].value,
        'new_name':sheet['B'+str(row)].value
        } for row in range(2,sheet.max_row + 1)]
    return lst_updates

def update_custom(swis_npm,props):
    print("\nStarting to Update Names..")
    for item in props:
        old_name = '{}'.format(item['Caption'])
        new_name = '{}'.format(item['new_name'])
        
        # Check to make sure there is an old and a new value
        if old_name == 'None':
            print('Skipping edit as existing name missing')
            continue
        elif new_name == 'None':
            print('Skipping edit as new name missing')
            continue
        else:
            try:
                results = swis_npm.query(
                    "SELECT Uri FROM Orion.Nodes WHERE Caption = @caption",
                     caption=old_name)  # Get Uri!
        
                print('Updating ' + old_name + ' to new name ' + new_name)
                uri = results['results'][0]['Uri']
                swis_npm.update(uri, Caption=new_name )
                continue
            except IndexError:
                print('The node {} to be updated does not exist'.format(item['Caption']))
                continue
                    
requests.packages.urllib3.disable_warnings()

    
if __name__ == '__main__':

    main()