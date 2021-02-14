def clean_file(switchname, tftp_path):
    
    with open(str(tftp_path) + switchname + "_int_status.txt", "rt") as sw_file:
        with open(str(tftp_path) + switchname + "_int_status_clean.csv", "w") as new_sw_clean_file:

            sw_lines = sw_file.readlines()

            line_number = 0
            for sw_line in sw_lines:
                i = 1
                index = 0
                word_count = len(sw_line.split())
                words = sw_line.split()
                Dict = eval(str(words))
                for each_word in words:
                    while i <= word_count: 
                        if line_number < 1:
                            i = word_count + 1
                            line_number = 2
                        else:
                            # If on 1st word add comma after word, this will be the interface.
                            if index == 0:
                                new_sw_clean_file.write(str(Dict[index]) + "\n")
                                i = word_count + 1