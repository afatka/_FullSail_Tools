#! python3

#written by Adam Fatka :: adam.fatka@gmail.com

#This script takes email addresses copied from excel (a1@2.com\ra2@2.com...)
#and adds a ', ' between them before putting them back on the clipboard. 

import pyperclip
orig_text = pyperclip.paste()
new_text = ','.join(orig_text.split('\r'))
pyperclip.copy(new_text)
print('conversion copied to clipboard')
