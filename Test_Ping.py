import sys
import os, fnmatch
import csv
import socket
from sys import platform
import tkinter as tk
from tkinter import filedialog


# Global Varialbles
file_name = ''
data = []
cur_dir = os.getcwd()
ping_timeout = "1"
ping_count = "2"

if platform == "win32":
    results_file = cur_dir + '\Ping_Results.txt'
else:
    results_file = cur_dir + '/Ping_Results.txt'


def main_menu(selected_file):
    global ping_count
    global ping_timeout

    if selected_file == '':
        selected_file = 'No File Selected Yet'
    else:
        selected_file = "Selected File: " + selected_file

    print('\n## Simple Ping Program ##\n')
    print('Enter 1 to Select a CSV file - ' + selected_file)
    #print('Enter 2 to Set Ping Options - Ping Count: ' + ping_count + " Timeout: " + ping_timeout + "s")
    print('Enter 2 to start Ping')
    print('\nEnter 0 to Exit the Program')
    value = input('\nPlease choose an option : ' )
    return value

def get_file():
    global cur_dir
    global file_name

    filetypes = "*.csv"
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

    option = main_menu(file_name)

    if option == "0":
        raise SystemExit()
    elif option == "1":
        file_name = get_file()
        #print(file_name)
        #raise SystemExit()
        if file_name == "":
            # input("\nNo file selected press enter to continue..")
            os.system('clear')
            main()
        else:
            data = open_file(file_name)
            #print(data)
            #input("\nPress enter to continue..")
            os.system('clear')
            main()
    # elif option == "2":
        
    #     os.system('clear')
    #     main()
    elif option == "2":
        if not data:
            input("No file selected, nothing to Ping!, press Enter Key")
            os.system('clear')
            main()
        else:
            num_replies = 0
            num_nonreplies = 0
            num_invalid = 0
            
            write_result = open(results_file, "w")
            ping_cmd = "ping "  # Set Default Ping Options

            if platform == "win32":
                ping_cmd = "ping -w 1000 "
            else:
                ping_cmd = "ping -q -c 2 -t 1 "
            for ip in data:
                ip_str = ''.join(ip)
                if is_valid_ipv4_address(ip_str) == True:
                    ping_result = (os.system(ping_cmd + ip_str))
                    if ping_result == 0:
                        write_result.write(ip_str + " Replies to Ping\n")
                    else:
                        write_result.write(ip_str + " Does not Reply\n")
                else:
                    print("\n" + ip_str + " is not a valid IP Address.")
                    write_result.write(ip_str + " is not a valid IP Address\n")

            write_result.close()
            
            final_result = open(results_file, "r")

            num_lines = len(final_result.readlines())
            final_result.close()

            num_replies = count_replies(results_file)
            num_nonreplies = count_nonreplies(results_file)
            num_invalid = count_invalid(results_file)

            print("\nNumber of IPs attempted: " + str(num_lines))
            print("Number IPs that Reply: " + str(num_replies))
            print("Number IPs that do not Reply: " + str(num_nonreplies))
            print("Number entries that are not a valid IP: " + str(num_invalid))
            input("\nPing finished, press Enter Key")
            os.system('clear')
            main()
    else:
        input("Not a valid option, press Enter Key")
        os.system('clear')
        main()


if __name__ == "__main__":
    main()