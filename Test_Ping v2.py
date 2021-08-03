import sys
import subprocess
import os, fnmatch
import csv
import socket
from sys import platform
import tkinter as tk
from tkinter import filedialog


# Define some Global Varialbles
file_name = ''
data = []
cur_dir = os.getcwd()

#  Check the operating system and set the working dir and clear command
if platform == "win32":
    results_file = cur_dir + '\\Ping_Results.txt'
    clear_cmd = 'cls'
else:
    results_file = cur_dir + '/Ping_Results.txt'
    clear_cmd = 'clear'


def main_menu(selected_file):

    if selected_file == '':
        selected_file = 'No File Selected Yet'
    else:
        selected_file = "Selected File: " + selected_file

    # Define some Menu Options
    print('\n### * Ping Program * Version 1.0 * Author:Kenny C * ###\n')
    print('Enter 1 to Select a CSV file - ' + selected_file)
    print('Enter 2 to start Ping')
    print('\nEnter 0 to Exit the Program')
    value = input('\nPlease choose an option : ' )
    return value

def get_file():
    global cur_dir
    global file_name

    root = tk.Tk()
    root.withdraw()
    selected_file = filedialog.askopenfilename(initialdir=cur_dir, 
        title="Open CSV file", 
        filetypes=(("CSV Files", "*.csv"),)
        )

    if not selected_file:
        return file_name
    else:
        return selected_file


def open_file(selected_file):
    with open(selected_file, newline='') as f:
        reader = csv.reader(f)
        data = list(reader)
    return data

def is_valid_ipv4_address(address):
    try:
        socket.inet_pton(socket.AF_INET, address)
    except AttributeError:  # no inet_pton here, sorry
        try:
            socket.inet_aton(address)
        except socket.error:
            return False
        return address.count('.') == 3
    except socket.error:  # not a valid address
        return False

    return True


def count_replies(filename):
    count = 0
    keyword = "Replies"
    with open(filename) as f:
        for line in f:
            if keyword in line:
                count+=1
        return count

def count_nonreplies(filename):
    count = 0
    keyword = "Does"
    with open(filename) as f:
        for line in f:
            if keyword in line:
                count+=1
        return count

def count_invalid(filename):
    count = 0
    keyword = "valid"
    with open(filename) as f:
        for line in f:
            if keyword in line:
                count+=1
        return count


def main():
    global file_name
    global data
    global results_file
    global clear_cmd

    option = main_menu(file_name)

    if option == "0":
        raise SystemExit()
    elif option == "1":
        file_name = get_file()
        #print(file_name)
        #raise SystemExit()
        if file_name == "":
            os.system(clear_cmd)  # Clear Screen
            main()  # Start Main function again to get main menu
        else:
            data = open_file(file_name)
            os.system(clear_cmd)  # Clear Screen
            main()  # Start Main function again to get main menu
    elif option == "2":
        if not data:
            input("No file selected, nothing to Ping!, press Enter Key")
            os.system(clear_cmd)  # Clear Screen
            main()  # Start Main function again to get main menu
        else:
            num_replies = 0
            num_nonreplies = 0
            num_invalid = 0
            num_lines = 0
            count_replies = 0
            count_nonreply = 0
            count_invalidIP = 0
            
            write_result = open(results_file, "w")
            #ping_cmd = "ping "  # Set Basic Ping Option

            # Check operating system and set Ping Options
            if platform == "win32":  # Windows Operating System
                ping_cmd = "ping -n 2 -w 100 "  # Set to 2 timeouts and 100 msec wait for response
            else:
                ping_cmd = "ping -q -W 100 -c 2 "

            
            # Start the Ping Loop
            for ip in data:
                num_lines +=1
                ip_str = ''.join(ip)  # Convert the list into a string variable
                if is_valid_ipv4_address(ip_str) == True:
                    ping_result = (os.system(ping_cmd + ip_str))
                    if ping_result == 0:
                        #write_result.write(ip_str + " Replies to Ping\n")  # Write output to file
                        count_replies +=1
                    else:
                        write_result.write("No Reply for IP Address : " + ip_str + "\n")  # Write output to file
                        count_nonreply +=1
                    try:
                        response = subprocess.check_output(
                            ['ping', '-q','-c', '3', '-W', '100', ip_str],
                            stderr=subprocess.STDOUT,  # get all output
                            universal_newlines=True  # return string not bytes
                        )
                        write_result.write(response + "\n")
                    except subprocess.CalledProcessError:
                        response = None
                else:
                    print("\n" + ip_str + " is not a valid IPv4 Address.")
                    write_result.write(ip_str + " is not a valid IPv4 Address\n")  # Write output to file
                    count_invalidIP +=1

            write_result.close()  # Close the output file
            
            final_result = open(results_file, "r")  # Open the final result file for reading
            #num_lines = len(final_result.readlines())  # Count the number of Lines
            final_result.close()  # Close File

            # Call Functions to count specific keywords
            # num_replies = count_replies(results_file)
            # num_nonreplies = count_nonreplies(results_file)
            # num_invalid = count_invalid(results_file)

            # Print the summary of the counts
            print("\nNumber of IPs attempted: " + str(num_lines))
            print("Number IPs that Reply: " + str(count_replies))
            print("Number IPs that do not Reply: " + str(count_nonreply))
            print("Number entries that are not a valid IP: " + str(count_invalidIP))
            print("Results file can be found in " + results_file)
            input("\nPing finished, press Enter Key")
            os.system(clear_cmd)  # Clear Screen
            main()  # Start Main function again to get main menu
    else:
        input("Not a valid option, press Enter Key")
        os.system(clear_cmd)
        main()  # Start Main function again to get main menu


if __name__ == "__main__":
    main()  # Call the main function in the Script 