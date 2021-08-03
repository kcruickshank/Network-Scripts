import sys
import os, fnmatch
import csv
import socket
import easygui
from sys import platform


def main_menu():
    print('\n## Simple Ping Program ##\n')
    print('Enter 1 to Select a CSV file')
    print('Enter 2 to Set Ping Options')
    print('Enter 3 to start Ping')
    value = input('\nPlease choose and option : ' )
    return value

def main():
    option = main_menu()

    if option == "1":




    print("You Selected Option " + str(option))
    input("\nPress Enter to continue..")
    os.system('clear')
    main()

if __name__ == "__main__":
    main()

