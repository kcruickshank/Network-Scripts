import netmiko
import getpass
import os
from os import name, system


# Global Variables
counter = 0

def main():
    authenticated = False
    while not authenticated:
        # Get the user logon details
        username, password = get_logon_details()
        clear_screen()
        # Check if authentication is ok
        authenticated = check_authentication(username,password)

    print('\nEnd of Test.')

def get_logon_details():
    username = input('Enter Username : ')
    password = getpass.getpass('Enter password : ')
    return username, password;

def check_authentication(username,password):
    global counter
    if username == 'Kenny' and password == 'cisco':
        print('Success your in!')
        return True
    else:
        if counter == 1:
            print("Authentication failed this is your last attempt make it count!\n")
            counter += 1
        elif counter > 1:
            print('\nAuthentication failed, program will now exit.')
            input('\nPress Enter key to exit')
            raise SystemExit
        else:    
            print('Authentication failed try again..\n')
            counter += 1

def clear_screen():
    if name == 'nt':
        _ = system('cls')
    else:
        _ = system('clear')

if __name__ == "__main__":
    
    main()