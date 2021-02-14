def clean_file(switchname, tftp_path):
    
    with open(str(tftp_path) + switchname + "_vlan_int.txt", "rt") as sw_file:
        with open(str(tftp_path) + switchname + "_new_file.csv", "w") as new_sw_clean_file:

            sw_lines = sw_file.readlines()

            line_number = 0
            for sw_line in sw_lines:
                i = 1
                index = 0
                word_count = len(sw_line.split())
                words = sw_line.split()
                last_word = word_count - 1  # Sets the last word index
                Dict = eval(str(words))
                for each_word in words:
                    while i <= word_count: 
                        # Start to check for Particular words
                        if str(Dict[index]) == "interface":
                            new_sw_clean_file.write(str(Dict[index]) + " " + str(Dict[i]) + " ")
                            print(str(Dict[index]) + " " + str(Dict[i]) + " ")
                            i += 1
                            index +=1
                        else:
                            i += 1
                            index +=1
