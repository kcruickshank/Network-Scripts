import pandas as pd
import re
from re import search
import os, signal, json
import shutil
import sys
import netmiko
from netmiko.ssh_exception import AuthenticationException, SSHException, NetMikoTimeoutException
from netmiko import ConnectHandler
import getpass
import datetime as datetime
from queue import Queue
from pprint import pprint
import threading


# These capture errors relating to hitting ctrl+C (I forget the source)
#signal.signal(signal.SIGPIPE, signal.SIG_DFL)  # IOError: Broken pipe
#signal.signal(signal.SIGINT, signal.SIG_DFL)  # KeyboardInterrupt: Ctrl-C

# Set the number of threads, I've found that 5 is the max
num_threads = 25

# Define the queue
enclosure_queue = Queue()

# Setup a print lock so only one thread prints at the one time
print_lock = threading.Lock()

# Define some variables to be used later in the script and ask user for some input
#username = input('Enter Username: ')
username = 'adm.cruickshank'
password = getpass.getpass('Enter Password: ')
loopback_int = input("What do you want to set the loopback interface source to? ")
ip_tftp_server = "10.251.6.35"
tftp_path_files = "\\" + "\\" + str(ip_tftp_server) + '\TFTP-Root\\Static_Files\\'
tftp_path = "\\" + "\\" + str(ip_tftp_server) + '\TFTP-Root\\IP_Helper_Files\\'
tftp_path_final = "\\" + "\\" + str(ip_tftp_server) + '\TFTP-Root\\IP_Helper_Files\\Final_Output_Files\\'
connection_details = []

# Define the current date and error log file info
now = datetime.datetime.now()
timestamp = now.strftime("%d-%m-%Y")  # Set timestamp to current system date
timestamp1 = now.strftime("%d/%m/%Y at %H:%M")  # Set timestamp to current system time including hour and minutes
logfile = timestamp + '_error_connection_log.txt'
tempoutput = '_temp_output.txt'


# Read the IP addresses from file
df_read_ip = pd.read_csv(str(tftp_path_files) + 'IP_Address_File.csv', header=None)


# Start to iterate through the IP's in the file

def deviceconnector(i,q):

    # Loop through the IP's
    while True:
        print("{}: Waiting for IP address...".format(i))
        ip_address = q.get()
        print("{}: Acquired IP: {}".format(i,ip_address))

        # Define a switch type
        switch = {
            "device_type": "cisco_ios",
                        "ip": ip_address,
                        "username": username,
                        "password": password,
        }

        # Test the ssh connection and handle any errors and output to text file
        try:
            net_connect = ConnectHandler(**switch)
        except (AuthenticationException):
            with print_lock:
                print("\n{}: ERROR: Authentication failed for {}. Stopping thread. \n".format(i,ip_address))
            error_log=open(tftp_path + str(logfile), "a")
            error_log.write("ERROR: Authentication failed for {} on the ".format(ip_address) + timestamp1 + "\n")
            error_log.close()
            # Add connection failed to list
            connection_details.append("Failed")
            q.task_done()
            continue 
        except (NetMikoTimeoutException):
            with print_lock:
                print("\n{}: ERROR: Connection to {} timed-out.\n".format(i,ip_address))
            error_log=open(tftp_path + str(logfile), "a")
            error_log.write("ERROR: Connection to {} timed-out on the ".format(ip_address) + timestamp1 + "\n")
            error_log.close()
            # Add connection failed to list
            connection_details.append("Failed")
            q.task_done()
            continue
        except (SSHException):
            with print_lock:
                print("\n{}: SSH might not be enable on: {} timed-out.\n".format(i,ip_address))
            error_log=open(tftp_path + str(logfile), "a")
            error_log.write("ERROR: SSH might not be enabled on: {} timed-out on the ".format(ip_address) + timestamp1 + "\n")
            error_log.close()
            # Add connection failed to list
            connection_details.append("Failed")
            q.task_done()
            continue
        except (EOFError):
            with print_lock:
                print("\n{}: End of liner error attempting device: {} timed-out.\n".format(i,ip_address))
            error_log=open(tftp_path + str(logfile), "a")
            error_log.write("ERROR: End of line error attempting device: {} timed-out on the ".format(ip_address) + timestamp1 + "\n")
            error_log.close()
            # Add connection failed to list
            connection_details.append("Failed")
            q.task_done()
            continue

        net_connect.enable()
        # Set tftp source interface    
        net_connect.send_config_set("ip tftp source-interface loopback " + str(loopback_int)) 

        # Save switch config
        net_connect.save_config()
        # Disconnect from the switch
        net_connect.disconnect

        # Set Task/Thread as complete
        q.task_done()

def main():

    # Setup the threads based on the number given above in the variables
    for i in range(num_threads):
        # Create the thread using the device connector as the function, pass in the thread number
        # and the queue object as the parameters
        thread = threading.Thread(target=deviceconnector, args=(i, enclosure_queue,))
        # Set thread up as a background job
        thread.setDaemon(True)
        # Start the thread
        thread.start()

    # Loop through the IP Address CSV and put the IP address into the queue
    for index, row in df_read_ip.iterrows():
        enclosure_queue.put(df_read_ip.iloc[index, 0])

    # Wait for all threads to be completed
    enclosure_queue.join()

    print('*** Source Interfaces Updated ***')

if __name__ == "__main__":
    
    # Call the main function
    main()