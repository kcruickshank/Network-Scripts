import easygui
import csv


msg='none'
title='Select a csv file'
filetypes=['*.csv']
default='*'

filename1= easygui.fileopenbox(msg,title,default,filetypes)


print(filename1)