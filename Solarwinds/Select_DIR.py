import re
import os
from os import path
import tkinter as tk
from tkinter import filedialog
import warnings
import getpass



def main():

    
    # Call function to writ new file
    write_file()

    # file = get_file()
    # print(file)
    
def write_file():
    # Check if directory Solarwinds exists on User C:\ Drive
    dir_path = r"C:\Solarwinds"
    file_path = r"C:\Solarwinds\Orion_Output.xlsx"

    if path.isfile(file_path):
        try:
            os.rename(file_path,file_path + "_")
            print("Access on file \"" + str(file_path) +"\" is available!")
            os.rename(file_path+"_",file_path)
        except OSError as e:
            message = "\nCan't create new file as " + str(file_path) + " Spreadsheet is Open, please close file and run again."
            print(message)
        

    if not path.isdir(dir_path):
        print(str(dir_path) + " does not exist")
        os.mkdir(dir_path)
        print(str(dir_path) + " Created!")
    
    
    
    
    
    
    # Set file outpu location to current user desktop
    #output_file_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')

    #new_file=open(str(output_file_path) + r"\KENNY_NEW.txt", "a")
    #new_file.write("Test")
    #new_file.close()

# def get_file():
#     # Create Dialogue file to select xlsx file
#     start_dir = os.getenv("USERPROFILE")
#     #start_dir = 'C:\TFTP-Root\Solarwinds'
#     #start_dir = 'C:\TFTP-Root\Solarwinds'
#     root = tk.Tk()  # Creates root window
#     root.withdraw()  # Hides the root window so only the file dialog shows
#     file_options = {'initialdir':start_dir, 'filetypes':[('Excel Files','.xlsx'),('All FIles','.*')]}
#     file_path = filedialog.askopenfilename(**file_options)  
#     if not file_path:
#         raise SystemExit  # If no file chosen exit the program
#     return file_path

if __name__ == '__main__':

    main()
    