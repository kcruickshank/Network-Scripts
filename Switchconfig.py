import netmiko
from netmiko import ConnectHandler
import getpass
import os
from os import name, system


# Global Variables
counter = 0

# Define a Cisco Switch Connection Handeller
cisco_switch = {
    'device_type': 'cisco_ios',
    'host': ip_address,
    'username': username,
    'password': password,
}


def main():
    authenticated = False
    # Loop the 2 functions to get the username and password 
    while not authenticated:
        # Get the user logon details
        username, password = get_logon_details()
        clear_screen()
        # Check if authentication is ok
        authenticated = check_authentication(username,password)

    print('\nEnd of Test and it works.')

def get_logon_details():
    username = input('Enter Username : ')
    password = getpass.getpass('Enter password : ')
    return username, password;

def check_authentication(username,password):
    global counter
    # Attempt an authentication to the network Device
    try:
        net_connect = ConnectHandler(**cisco_switch)
        print('Success your in!')
        return True
    except netmiko.NetMikoAuthenticationException:
        # Check Counter value to allow 3 attempts to authenticate before exiting the program
        if counter == 0:
            print('Authentication failed try again..\n')
            counter += 1
        if counter == 1:
            print("Authentication failed this is your last attempt make it count!\n")
            counter += 1
        elif counter > 1:
            print('\nAuthentication failed, program will now exit.')
            input('\nPress Enter key to exit')
            raise SystemExit  
            

def clear_screen():
    if name == 'nt':
        _ = system('cls')
    else:
        _ = system('clear')

if __name__ == "__main__":
    
    main()