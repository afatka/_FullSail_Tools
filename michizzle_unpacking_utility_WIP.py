#! usr/bin/env python3
# Michizzle_Unpack_Utility
# Input: 
#	Full Sail LMS Batch Download Zip file(s)
# Output:
#	Create a folder in input location, named the same as the zip
#	Each student gets a folder of format "Lastname_Firstname_StudentIDNumber"
#	Place all student content into their folder, expand all content

# Created by:
# Adam Fatka
# adam.fatka@gmail.com


'''
Creation notes:
		Use Qt for GUI feedback. No output to the terminal
		Design to apply as a service (RMB)


TODO: Create Output Window
TODO: Notify User of start
TODO: Validate Input
TODO: For each input: Do
TODO: 	- Create parent folder
TODO: 	- Create Student Folder
TODO: 	-  Unpack Student zip recursively, if necessary
TODO: 	-  - If student contents folder already exists, iterate. !!!




'''









### Old Michizzle is below ###
import zipfile, os, sys, io, time

def do_work(*args):

	version_number = '2'


	print('{} - Version: {}'.format('Michizzle Unpacking Utility', version_number))
	input_files = validate_input(args)
	for file in input_files:
		print('Unpacking zip: {}'.format(os.path.basename(file)), end = '', flush = True)
		unpack(file)
	print('\nMichizzle Unpacking Complete.')
		

def print_usage():
	print('''
	Unpack_fszip Usage:
	In the terminal type: python3<space>
	Drag the unpack_fszip script file into the terminal
	Drag 1 or more zip files into the terminal
	Press <enter>
	*Magic!*
	''')

def validate_input(sys_args):
	input_arg = sys_args[0]

	if len(input_arg) == 1:
		print_usage()
		sys.exit()

	# iterate over input files 
	# verify that they; Are files, are zip files
	input_files = input_arg[1:]
	for arg in input_files:
		if not arg.endswith('.zip'):
			raise ValueError('Incorrect input format. File must be of extension ".zip".')
		if not os.path.isfile(input_arg[1]):
			raise ValueError('Incorrect input. Input must be a file')
		if not zipfile.is_zipfile(input_arg[1]):
			raise ValueError('Incorrect input. File must be zip file')
	return input_files

def unpack(input_file):

	path = os.path.splitext(input_file)[0]
	path = validate_path(path)
	deal_with_zip(path, input_file, True)
	print(' - Unpacking complete')

def validate_path(path):
	i = 1
	orig_path = path
	while True:
		if not os.path.exists(path):
			return path
		path = '{} {}'.format(orig_path, i)
		i+= 1

def deal_with_zip(path, zip_file, split):
	with zipfile.ZipFile(zip_file, 'r') as z:
		for f in z.infolist():
			name = f.filename
			if not name.endswith((os.sep, '.log', '.DS_Store')) and not name.startswith(('__')):
				if name.endswith('.zip'):
					student_zip = io.BytesIO(z.read(name))
					try:
						deal_with_zip(path, student_zip, False)
					except Exception as e:
						print('\n{} cannot be extracted. Possibly corrupt. {}'.format(name, e))
				else: 
					filename_with_path = file_from_zip_to_new_dir(path, name, split)
					write_files(filename_with_path, z.read(name), f.date_time)

def stop():
	raise ValueError('STAHP!')

def file_from_zip_to_new_dir(path, fromZip, split):
	# print('from zip')
	if split == True: 
		newFile = os.path.join(path, fromZip.split(os.sep, 2)[-1])
		# print('spliting: {}'.format(fromZip))
	else:
		newFile = os.path.join(path, fromZip)
	# print('old: {}'.format(fromZip))
	# print('New: {}'.format(newFile))
	return newFile


#maintaining timestamp from http://stackoverflow.com/a/9813471/5079170
def write_files(file_nameWPath, file_to_write, file_datetime):
	# print('write file')
	# print('file_nameWPath:{}'.format(file_nameWPath))
	os.makedirs(os.path.dirname(file_nameWPath), exist_ok = True)
	filename = validate_filenameWPath(file_nameWPath)
	with open(filename, 'wb') as f:
		f.write(file_to_write)
	date_time = time.mktime(file_datetime + (0, 0, -1))
	os.utime(filename, (date_time, date_time))

def validate_filenameWPath(file_nameWPath):
	# print('validate_filenameWPath')
	i = 1
	ext_split = os.path.splitext(file_nameWPath)
	while True:
		if not os.path.isfile(file_nameWPath):
			return file_nameWPath
		file_nameWPath = '{} {}{}'.format(ext_split[0], i, ext_split[1])
		i += 1

def log(message):
	''' Print log. Can be shut off'''
	do_print = True
	if do_print:
		print(message)

# If ran as main... do it
if __name__ == '__main__':
	do_work(sys.argv)




