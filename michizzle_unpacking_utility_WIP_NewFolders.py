#! usr/bin/env python3
#unpack_fszip.py / Muh-shizzle (Michizzle?) 
#take a zip file from FSO portal and unpack all the zips down to the student content.

# Created by:
# Adam Fatka
# adam.fatka@gmail.com

import zipfile, os, sys, io, time

def do_work(*args):

	version_number = '1.5'


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
	print('\n')
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

###
### TODO ### Add some sort of "New folder" identifier. This will get passed to write_files
###

###
### This may be able to be handled with a context manager instead....
### each time the context is entered check the path for exists and iterate... ?
###

def deal_with_zip(path, zip_file, split):
	with zipfile.ZipFile(zip_file, 'r') as z:
		for f in z.infolist():
			name = f.filename
			if not name.endswith((os.sep, '.log', '.DS_Store')) and not name.startswith(('__')):
				print('Name: {}'.format(name))
				if name.endswith('.zip'):
					student_zip = io.BytesIO(z.read(name))
					try:
						deal_with_zip(path, student_zip, False)
					except Exception as e:
						print('\n{} cannot be extracted. Possibly corrupt. {}'.format(name, e))
				else: 
					filename_with_path = file_from_zip_to_new_dir(path, name, split)
					# write_files(filename_with_path, z.read(name), f.date_time)

		
def stop():
	raise ValueError('STAHP!')

def file_from_zip_to_new_dir(path, fromZip, split):
	if split == True: 
		newFile = os.path.join(path, fromZip.split(os.sep, 2)[-1])
	else:
		newFile = os.path.join(path, fromZip)
	return newFile

###
### TODO ### If new folder, iteratate file_nameWPath until it doesn't exist
###

#maintaining timestamp from http://stackoverflow.com/a/9813471/5079170
def write_files(file_nameWPath, file_to_write, file_datetime):
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



if __name__ == '__main__':
	# print("args: {}".format(sys.argv))
	print('\a')
	do_work(sys.argv)