#! env/bin/python3

# collect grade info docs text, format for excel paste

import sys, os, re
import pyperclip # external dependency

def fileFinderOSWalk( directory_to_walk, *args):
	file_types = ('.txt')
	files_of_interest = []

	directoryDepth = 0
	max_recursion = 8

	if not os.path.isdir(directory_to_walk):
		raise ValueError('Input is not a directory')

	for directoryName, subDirectoryName, fileList in os.walk(directory_to_walk):
		if directoryDepth == 0:
			directoryDepth = len(directoryName.split(os.sep))
		if len(directoryName.split(os.sep)) - directoryDepth >= max_recursion:
			print('os.walk recursion break triggered')
			break

		# xls_files =[]
		for item in fileList:
			if item.endswith(file_types):
				files_of_interest.append(os.path.join(directoryName, item))

	return files_of_interest

text_files = fileFinderOSWalk(sys.argv[1])

text_to_copy = ''
for tf in text_files:
	with open(tf) as f:
		file_string = ''
		# student_info = ''
		for line in f:

			escape_line = re.sub('"', r"'", line)
			file_string += escape_line
		grade_total = re.search(r'Grade Total:\d*%<br>', file_string)

		if grade_total:
			text_to_copy += '{}\t'.format(grade_total.group(0).split('%')[0].split(':')[1])
			#
			#
			#Included '\n' instead of newline... escape only " "
			#
			#

			# text_to_copy += '"{}"\n'.format(repr(file_string))

			# text_to_copy += '"{}"\n'.format(re.sub('"',r'\"' ,file_string))
			text_to_copy += '"{}"\n'.format(file_string)
			# text_to_copy += '"{}"\n'.format(re.sub(r'"',r'\"',file_string))

pyperclip.copy(text_to_copy)
print('content copied')

		

