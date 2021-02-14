import os


def menu():
    print("\n**Get Solarwind Nodes Program **")
    print("\nPress 1 to get new nodes and create new spreadsheet")
    print("Press 2 to exit the program")
    choice = input("\nEnter Choice: ")
    return choice

# choice = menu()

# print(choice)


loop = True

while loop == True:
    choice = menu()
    
    if choice == 1:
        print("Run Program....")
    elif choice == 2:
        print("Program Quit")
        break
        raise SystemExit
    else:
        print("Invalind answer, accepts only number 1 or 2\n")
        input('Press Enter key...')
        os.system('cls')
