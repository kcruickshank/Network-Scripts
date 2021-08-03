import re


user_input = input("Enter Value: ")
mac_pattern = "[a,b,c,d,e,f,A,B,C,D,E,F0-9]+\.[a,b,c,d,e,f,A,B,C,D,E,F0-9]+\.[a,b,c,d,e,f,A,B,C,D,E,F0-9]"
num_pattern = "^[ 0-9]+$"
int_pattern = "^[Gi,Fa,Twe,Te,Hu]"
test_pattern = "dynamic"


if re.search(test_pattern,user_input):
    print('User entered valid number: ' + str(user_input))
else:
    print('Not Matched!')


# if re.search(int_pattern,user_input):
#     print('Found')
# else:
#     print('Not found')


