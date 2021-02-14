import re
import requests
from orionsdk import SwisClient
import openpyxl
import os
from os import path
import warnings
import getpass
import datetime as datetime


# Define some Variables for DNS Updates
dns_zone = "cns.muellergroup.com"
dns_server_1 = "10.96.65.100"
dns_server_2 = '10.96.65.101'
dns_server_3 = '10.80.50.45'


# Log file and time variables
now = datetime.datetime.now()  # Get current time
timestamp = now.strftime("%d-%m-%Y")  # Set timestamp to current system date
timestamp_1 = now.strftime("- logged at %H:%M")  # Set timestamp to current system date include hours and minutes
dns_updt_logfile = r"C:\Solarwinds\\DNS Update Error Log " + str(timestamp) + ".txt"
dns_del_logfile = r"C:\Solarwinds\\DNS Delete Error Log " + str(timestamp) + ".txt"
custom_updt_logfile = r"C:\Solarwinds\\Node Name Update Error Log " + str(timestamp) + ".txt"


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
        input('\nPress Enter key to return to menu')
        raise SystemExit

    #print('\n** Use your normal Active Directory Account to connect to Solarwinds **\n')
    
    verify = False
    if not verify:
        from requests.packages.urllib3.exceptions import InsecureRequestWarning
        requests.packages.urllib3.disable_warnings(InsecureRequestWarning)
    
    while not authenticated:
        #username = input('Enter Username: ')
        #password = getpass.getpass('Enter Password: ')
        username = "mapsacc"
        password = "nMUDIXKZ3V"
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
    
    print("\n*** Node Name and DNS Update Script Started ***")

    file = file_path
    node_details = read_file(file)

    if len(node_details) == 0:
        print("\nThere are no details to update")
        input('\nPress enter key to Exit')
        #raise SystemExit
    else:
        custom_code = update_node(swis_npm,node_details)
        dns_delete_code = delete_dns(swis_npm,node_details)
        dns_update_code = update_dns(swis_npm,node_details)

        # Print a final Summary of the program output
        print("\n*** Summary ***")

        # Check the Custom Attributes Update Code
        if custom_code == 1:
            print("\n- There were errors updating node name, please review error log file {}".format(custom_updt_logfile))
        else:
            print("\n- Node name updates completed with no errors.")

        # Check the DNS Delete Code
        if dns_delete_code == 1:
            print("- There were errors deleting some DNS entries, please review error log file {}".format(dns_del_logfile))
        else:
            print("- DNS deletions completed with no errors.")

        # Check the DNS Update Code
        if dns_update_code == 1:
            print("- There were errors updating some DNS entries, please review error log file {}".format(dns_updt_logfile))
        else:
            print("- DNS updates completed with no errors.")

        input('\nPress enter key to Exit')

# Code commented out that that provided pop up box to select file
# was too long to select file everytime when it did not change
# So not used anymore
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
        'ip_add':sheet['B'+str(row)].value,
        'dns_name':sheet['C'+str(row)].value,
        'new_name':sheet['D'+str(row)].value
        } for row in range(2,sheet.max_row + 1)]
    return lst_updates

def update_node(swis_npm,props):
    custom_code = 0
    print("\nStarting Solarwind Node Name Updates..")
    for item in props:
        old_name = '{}'.format(item['Caption'])
        new_name = '{}'.format(item['new_name'])
        ip_add = '{}'.format(item['ip_add'])
        
        # Check to make sure there is an old and a new value to update node name
        if old_name == 'None':
            error_log=open(str(custom_updt_logfile), "a")
            error_log.write("ERROR: Existing node name not set in spreadsheet for IP Address {} ".format(ip_add) + str(timestamp_1) + "\n")
            error_log.close()
            custom_code = 1
            continue
        elif new_name == 'None':
            error_log=open(str(custom_updt_logfile), "a")
            error_log.write("ERROR: New node name not set in spreadsheet for IP Address {} ".format(ip_add) + str(timestamp_1) + "\n")
            error_log.close()
            custom_code = 1
            continue
        else:
            try:
                # Start by updating the Node Name
                # Look up the node based on the name in the existing name column
                results = swis_npm.query(
                    "SELECT Uri FROM Orion.Nodes WHERE Caption = @caption",
                     caption=old_name)  # Get Uri!
                
                # Print output of whats happening to the screen
                print('- Updating Node: ' + old_name + ' to new name: ' + new_name)
                uri = results['results'][0]['Uri']
                swis_npm.update(uri, Caption=new_name )         
                
                # Edit the DNS and System Name Fields
                swis_npm.update(uri, DNS=new_name + "." + dns_zone)
                swis_npm.update(uri, SysName=new_name + "." + dns_zone)
            except IndexError:
                error_log=open(str(custom_updt_logfile), "a")
                error_log.write("ERROR: The existing node name {} trying to be updated does not exist in Solarwinds ".format(old_name) + str(timestamp_1) + "\n")
                error_log.close()
                custom_code = 1
                #print('The node {} to be updated does not exist'.format(item['Caption']))
                continue   
    return custom_code

def delete_dns(swis_npm,props):
    dns_delete_code = 0
    print("\nStarting Solarwinds IPAM DNS Deletion..")
    for item in props:
        ip_add = '{}'.format(item['ip_add'])
        old_dns = '{}'.format(item['dns_name'])
        
        # Check to make sure the required values are in the update sheet to update DNS
        if old_dns == 'None':
            error_log=open(str(dns_del_logfile), "a")
            error_log.write("ERROR: DNS Name is blank in Edit Node Name sheet for IP Address {}, cannot delete DNS entry ".format(ip_add) + str(timestamp_1) + "\n")
            error_log.close()
            dns_delete_code = 1
            continue
        elif ip_add == 'None':
            error_log=open(str(dns_del_logfile), "a")
            error_log.write("ERROR: IP Address is blank in Edit Node Name sheet for DNS Name {}, cannot delete DNS entry ".format(old_dns) + str(timestamp_1) + "\n")
            error_log.close()
            dns_delete_code = 1
        else:
            # First try to remove existing DNS Records
            print("- Starting to remove DNS record " + str(old_dns))
            try:
                swis_npm.invoke('IPAM.IPAddressManagement', 'RemoveDnsARecord',old_dns,ip_add,dns_server_1,dns_zone)
                swis_npm.invoke('IPAM.IPAddressManagement', 'RemoveDnsARecord',old_dns,ip_add,dns_server_2,dns_zone)
                swis_npm.invoke('IPAM.IPAddressManagement', 'RemoveDnsARecord',old_dns,ip_add,dns_server_3,dns_zone)
            except requests.exceptions.HTTPError:
                error_log=open(str(dns_del_logfile), "a")
                error_log.write("ERROR: DNS entry {} does not exist for IP Address {} ".format(old_dns,ip_add) + str(timestamp_1) + "\n")
                error_log.close()
                dns_delete_code = 1
                continue

            print("- Finished removing DNS record " + str(old_dns))
    return dns_delete_code

def update_dns(swis_npm,props):
    dns_update_code = 0
    print("\nStarting Solarwinds IPAM DNS Updates...")
    for item in props:
        ip_add = '{}'.format(item['ip_add'])
        new_name = '{}'.format(item['new_name'])
        
        # Check to make sure the required values are in the update sheet to update DNS
        if new_name == 'None':
            error_log=open(str(dns_updt_logfile), "a")
            error_log.write("ERROR: New Name is blank in Edit Node Name sheet for IP Address {}, cannot update DNS entry ".format(ip_add) + str(timestamp_1) + "\n")
            error_log.close()
            dns_update_code = 1
            continue
        elif ip_add == 'None':
            error_log=open(str(dns_updt_logfile), "a")
            error_log.write("ERROR: IP Address is blank in Edit Node Name sheet for New DNS Name {}, cannot update DNS entry ".format(new_name) + str(timestamp_1) + "\n")
            error_log.close()
            dns_update_code = 1
            continue
        else:
            results_dns = swis_npm.query("SELECT Uri, Status FROM IPAM.IPNode WHERE IPAddress = " + "'" + ip_add + "'")
            
            # Check if results are empty and write to error file that IP does not exist
            if len(results_dns['results']) == 0:
                error_log=open(str(dns_updt_logfile), "a")
                error_log.write("ERROR: IP Address {} does not exist in IPAM, cannot update DNS ".format(ip_add) + str(timestamp_1) + "\n")
                error_log.close()
                dns_update_code = 1
                continue
            else:
                uri_dns = results_dns['results'][0]['Uri']
                status = results_dns['results'][0]['Status']

                if status == 2:
                    error_log=open(str(dns_updt_logfile), "a")
                    error_log.write("ERROR: IP Address {} has status available in IPAM, cannot update attributes ".format(ip_add) + str(timestamp_1) + "\n")
                    error_log.close()
                    dns_update_code = 1
                else:  
                    #Add Hostname field in IPAM IP Address
                    swis_npm.update(uri_dns, DnsBackward=new_name + "." + dns_zone)
                    #Try to add new DNS Records
                    print("- Starting to add DNS record " + str(new_name) + "." + str(dns_zone))
                    try:
                        swis_npm.invoke('IPAM.IPAddressManagement', 'AddDnsARecord',new_name,ip_add,dns_server_1,dns_zone)
                        swis_npm.invoke('IPAM.IPAddressManagement', 'AddDnsARecord',new_name,ip_add,dns_server_2,dns_zone)
                        swis_npm.invoke('IPAM.IPAddressManagement', 'AddDnsARecord',new_name,ip_add,dns_server_3,dns_zone)
                    except requests.exceptions.HTTPError:
                        error_log=open(str(dns_updt_logfile), "a")
                        error_log.write("ERROR: Unable to update DNS record for IP Address {} ".format(ip_add) + str(timestamp_1) + "\n")
                        error_log.close()
                        dns_update_code = 1
                        print("Error adding DNS record!")
                        continue
                
                    print("- Finished adding DNS record " + str(new_name) + "." + str(dns_zone))
    
    return dns_update_code
            


requests.packages.urllib3.disable_warnings()


if __name__ == '__main__':

    main()