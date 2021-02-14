
def clean_file(switchname, tftp_path):

    with open(str(tftp_path) + switchname + "_int_status.txt", "rt") as int_status_file:
        with open(str(tftp_path) + switchname + "_int_status_clean.csv", "w") as new_int_status_file:

            int_status_lines = int_status_file.readlines()
            max_word_count = 0

            # The first for loop has to go through the lines and find the most amount of words. This will
            # then be used to set or mac number of columns
            for status_line in int_status_lines:
                word_count = len(status_line.split())
                if max_word_count < word_count:
                    max_word_count = word_count

            # This code takes the max number of words and creates fake heading up to the max number of words
            i = 1
            index = 0
            last_word = max_word_count -1
            while i <= max_word_count:
                if index < last_word:
                    new_int_status_file.write("Heading" + str(index) + ",")
                    i += 1
                    index += 1
                else:
                    new_int_status_file.write("Heading" + str(index) + "\n")
                    i += 1
                    index += 1

            line_number = 0
            for status_line in int_status_lines:
                i = 1
                j = 0
                index = 0
                word_count = len(status_line.split())
                words = status_line.split()
                last_word = word_count - 1  # Sets the last word index
                Dict = eval(str(words))
                # print("There are " + str(word_count) + " words in this line and they are:")
                for each_word in words:
                    while i <= word_count:
                        if line_number < 1:
                            # This part write all the words in on the top line of the text file on new line
                            if index < last_word:
                                new_int_status_file.write(str(Dict[index]) + ",")
                                i += 1
                                index += 1
                            else:
                                new_int_status_file.write(str(Dict[index]) + "\n")
                                i += 1
                                index += 1
                                line_number += 1
                        else:
                            # This part then loops through all the interface lines and writes out the comma separated
                            # words
                            if index == 0:
                                if word_count < 2:
                                    new_int_status_file.write(str(Dict[index] + '\n'))
                                    i += 1
                                    index += 1
                                else:
                                    new_int_status_file.write(str(Dict[index] + ",Blank,"))
                                    i += 1
                                    index += 1
                            else:
                                if str(Dict[index]) == "connected":
                                    #new_int_status_file.write(" ,")
                                    new_int_status_file.write(str(Dict[index] + ","))
                                    j = index
                                    i += 1
                                    index += 1
                                elif str(Dict[index]) == "notconnect":
                                    # print("Found the word " + str(Dict[index]) + " and is at index " + str(index))
                                    #new_int_status_file.write(" ,")
                                    new_int_status_file.write(str(Dict[index] + ","))
                                    j = index
                                    i += 1
                                    index += 1
                                elif str(Dict[index]) == "err-disabled":
                                    # print("Found the word " + str(Dict[index]) + " and is at index " + str(index))
                                    #new_int_status_file.write(" ,")
                                    new_int_status_file.write(str(Dict[index] + ","))
                                    j = index
                                    i += 1
                                    index += 1
                                elif str(Dict[index]) == "disabled":
                                    # print("Found the word " + str(Dict[index]) + " and is at index " + str(index))
                                    #new_int_status_file.write(" ,")
                                    new_int_status_file.write(str(Dict[index] + ","))
                                    j = index
                                    i += 1
                                    index += 1
                                else:
                                    if j > 0:
                                        if index < last_word:
                                            new_int_status_file.write(str(Dict[index]) + ",")
                                            i += 1
                                            index += 1
                                        else:
                                            new_int_status_file.write(str(Dict[index]) + "\n")
                                            i += 1
                                    else:
                                        i += 1
                                        index += 1

