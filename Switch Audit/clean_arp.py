def clean_file(switchname, tftp_path):

    with open(str(tftp_path) + switchname + "_arp.txt", "rt") as read_file:
        with open(str(tftp_path) + switchname + "_arp_clean.csv", "w") as write_file:

            file_lines = read_file.readlines()

            line_number = 0

            for line in file_lines:
                i = 1
                index = 0
                word_count = len(line.split())
                words = line.split()
                last_word = word_count - 1  # Sets the last word index
                Dict = eval(str(words))
                for each_word in words:
                    while i <= word_count:
                        if line_number < 1:
                            # This part will skip the top line
                            if index < last_word:
                                i += 1
                                index += 1
                            else:
                                i += 1
                                index += 1
                                line_number += 1
                        else:
                            if index < last_word:
                                write_file.write(str(Dict[index]) + ',')
                                i += 1
                                index += 1
                            else:
                                write_file.write(str(Dict[index]) + '\n')
                                i += 1
                                index += 1 