def clean_file(ip_address, tftp_path):
    
    with open(str(tftp_path) + ip_address + "_Wlan_Summary.txt", "rt") as file:
        with open(str(tftp_path) + ip_address + "_Wlan_Summary_clean.csv", "w") as new_clean_file:

            lines = file.readlines()

            line_number = 0
            for line in lines:
                i = 1
                index = 0
                word_count = len(line.split())
                words = line.split()
                Dict = eval(str(words))
                for each_word in words:
                    while i <= word_count: 
                        if line_number < 1:
                            i = word_count + 1
                            line_number = 2
                        else:
                            # If on 1st word add comma after word, this will be the interface.
                            if index == 0:
                                new_clean_file.write(str(Dict[index]) + "\n")
                                i = word_count + 1