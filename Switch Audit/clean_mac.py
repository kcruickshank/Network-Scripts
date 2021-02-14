def clean_file(switchname, tftp_path):

    with open(str(tftp_path) + switchname + "_mac_add_temp1.txt", "rt") as mac_add_file:
        with open(str(tftp_path) + switchname + "_mac_add_clean.csv", "w") as new_mac_add_file:

            mac_add_lines = mac_add_file.readlines()

            for status_line in mac_add_lines:
                i = 1
                index = 0
                word_count = len(status_line.split())
                words = status_line.split()
                last_word = word_count - 1  # Sets the last word index
                Dict = eval(str(words))
                for each_word in words:
                    while i <= word_count:
                        if index < last_word:
                            new_mac_add_file.write(str(Dict[index]) + ',')
                            i += 1
                            index += 1
                        else:
                            new_mac_add_file.write(str(Dict[index]) + '\n')
                            i += 1
                            index += 1