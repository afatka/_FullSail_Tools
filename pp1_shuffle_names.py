#! #! env/bin/python3

import sys
import pyperclip #external dependency

def shuffle_names(number_of_groups):
	names = pyperclip.paste().split('\r')
	return_string = ''
	i = 1
	for name in names:
		if name != '':
			if i == int(number_of_groups):
				return_string += '{}{}'.format(name, '\r')
				i = 1
			else:
				return_string += '{}{}'.format(name, '\t')
				i += 1
	# print('{}'.format(repr(return_string)))
	pyperclip.copy(return_string)
	print('Names Shuffled')

def print_usage():
	print('''
		pp1_shuffle_names Usage:
		Copy a column of names. 
		launch script with argument number_of_groups.
		paste shuffled names
		''')

if len(sys.argv) ==1:
	print_usage()
else:
	shuffle_names(sys.argv[1])