import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
from pandas import DataFrame
import numpy as np
import re
from re import search
import os
import shutil
import sys
import netmiko
from netmiko.ssh_exception import AuthenticationException, SSHException, NetMikoTimeoutException
from netmiko import ConnectHandler
import getpass
import clean_int_brief as clean_brief
import datetime as datetime


# Define some variable to be used later in the script and ask user for some input
#username = input('Enter Username: ')
ip_tftp_server = "10.251.6.35"
tftp_path_files = "\\" + "\\" + str(ip_tftp_server) + '\TFTP-Root\\Static_Files\\'
tftp_path = "\\" + "\\" + str(ip_tftp_server) + '\TFTP-Root\\IP_Helper_Files\\'
tftp_path_final = "\\" + "\\" + str(ip_tftp_server) + '\TFTP-Root\\IP_Helper_Files\\Final_Output_Files\\'
tftp_path_move_old = "\\" + "\\" + str(ip_tftp_server) + '\TFTP-Root\\IP_Helper_Files\\Old_Text_Files\\'
log_file = 'helper_check_error_log.txt'

df0 = pd.DataFrame()
filenames = os.listdir(tftp_path)
os.chdir(tftp_path)
for filename in filenames:  # Convert the text file names to lowercase
    if '.txt' in filename:
        newfilename = filename.lower()
        os.rename(filename, newfilename)
df0['FILES'] = os.listdir(tftp_path)
df0 = df0.loc[~df0['FILES'].str.contains('Files', flags=re.I, regex=True)]
df0 = df0.loc[~df0['FILES'].str.contains('.xlsx', flags=re.I, regex=True)]
df0 = df0.loc[~df0['FILES'].str.contains('.csv', flags=re.I, regex=True)]
full_list_files = df0['FILES'].tolist()
df0 = df0.drop_duplicates()
df0.reset_index(drop=True, inplace=True)
file_list = df0['FILES'].tolist()

print(file_list)
