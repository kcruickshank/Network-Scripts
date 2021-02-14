def clean_file(switchname, tftp_path):

    with open(str(tftp_path) + switchname + "_int_desc.txt", "rt") as int_desc_file:
        with open(str(tftp_path) + switchname + "_int_desc_clean.csv", "w") as new_int_desc_file:

            int_desc_lines = int_desc_file.readlines()
            max_word_count = 0
            string1 = "admin"
            string2 = "down"

            # The first for loop has to go through the lines and find the most amount of words. This will
            # then be used to set or max number of columns
            for desc_line in int_desc_lines:
                word_count = len(desc_line.split())
                if max_word_count < word_count:
                    max_word_count = word_count
            # This code takes the max number of words and creates fake heading up to the max number of words
            i = 1
            index = 0
            last_word = max_word_count - 1
            while i <= max_word_count:
                if index < last_word:
                    new_int_desc_file.write("Heading" + str(index) + ",")
                    i += 1
                    index += 1
                else:
                    new_int_desc_file.write("Heading" + str(index) + "\n")
                    i += 1
                    index += 1

            for desc_line in int_desc_lines:
                i = 1
                index = 0
                word_count = len(desc_line.split())
                words = desc_line.split()
                last_word = word_count - 1  # Sets the last word index
                Dict = eval(str(words))
                for each_word in words:
                    while i <= word_count:
                        if index < last_word:  
                            # If on 1st word add comma after word, this will be the interface.
                            if index == 0:
                                new_int_desc_file.write(str(Dict[index]) + ",")
                                i += 1
                                index += 1
                            # Check word 2, if it contains the word admin, then check the if the next word contains
                            # down, if it does then we need to put these in the same column index of 1
                            elif index == 1:
                                if str(Dict[index]).lower == string1.lower:
                                    if str(Dict[i]).lower == string2.lower:
                                        new_int_desc_file.write(str(Dict[index]) + " " + str(Dict[i]) + ",")
                                        i += 2
                                        index +=2
                                else:
                                    new_int_desc_file.write(str(Dict[index]) + ",")
                                    i += 1
                                    index += 1
                            else:
                                new_int_desc_file.write(str(Dict[index]) + ",")
                                i += 1
                                index += 1
                        else:
                            new_int_desc_file.write(str(Dict[index]) + "\n")
                            i += 1