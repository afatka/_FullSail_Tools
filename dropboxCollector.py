#! #!/usr/bin/env python3

# Dropbox grade Collector
# Dependencies::
# xlrd, pyperclip

import os, sys
import xlrd, pyperclip #these are dependencies

class DropBoxValueCollector(object):
	'Collect information from dropbox files'

	def __init__(self, *args, **kwargs):
		self.status('Starting Data Collection...')
		self.development = kwargs.get('development', False)
		self.file_types = kwargs.get('files', ('.xls')) #file types to find
		self.cell_and_title_pairs = kwargs.get('cells_and_titles', 
		[['week 1', [9, 11]],
		 ['week 2', [13, 11]], 
		 ['week 3', [17, 11]],
		 ['Final', [29, 24]]]) #cell [row, col] ##INDEX 0 for BOTH ROW AND COLUMN!!! 
		self.sheet_of_interest = kwargs.get('sheet', 'Instructor')
		self.files_of_interest = []
		self.students_info = []

		if len(args[0]) != 2:
			self.log('args: {}'.format(args[0]))
			raise ValueError('Input a single directory.')

		# TODO - crawl provided folder and find xls files
		self.fileFinderOSWalk(args[0][1])

		# TODO - iterate through xls files and collect grade data and student names
		self.collectFileInfo()

		# TODO - create string to paste into spreadsheet 
		self.create_string_to_paste()
		print('\n')

	def create_string_to_paste(self):
		self.status('Creating String')
		final_string = 'Student\t'
		for t in self.cell_and_title_pairs:
			final_string += t[0] + '\t'
		final_string += '\n'
		for student in self.students_info:
			final_string += student[0] + '\t' 
			for v in student[1]:
				final_string += str(v) + '\t'
			final_string += '\n'

		self.log('\n{}'.format(final_string))
		pyperclip.copy(final_string)
		self.status('Info Copied.')

	def collectFileInfo(self):
		self.status('Gathering Grades')

		# step = int(len(self.files_of_interest)/10)
		iter_count = 1

		for f in self.files_of_interest:

			# if iter_count%step == 0:
			self.progress(iter_count)
			iter_count += 1

			wb = xlrd.open_workbook(f)
			for sheet_name in wb.sheet_names():
				if sheet_name == self.sheet_of_interest:
					ws = wb.sheet_by_name(sheet_name)

					student = []
					grade_info = []
					self.log('f: {}'.format(f))
					student.append(os.path.basename(os.path.split(f)[0]))
					for c in self.cell_and_title_pairs:
						grade_info.append(ws.cell(c[1][0],c[1][1]).value)
					student.append(grade_info)
					self.students_info.append(student)

		self.log(self.students_info)
		self.status('Grades Gathered')

	def fileFinderOSWalk(self, directory_to_walk, *args):
		self.status('Collecting Files')

		directoryDepth = 0
		max_recursion = 6

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
				if item.endswith(self.file_types):
					self.log('directory name is: {}'.format(directoryName))
					self.log('file name: {}'.format(item))
					self.files_of_interest.append(os.path.join(directoryName, item))
					# xls_files.append(os.path.join(directoryName, item))

			# append_file = sorted(xls_files, key=lambda x: os.stat(x).st_mtime)	
			# self.files_of_interest.append(append_file[0])

		self.status('Files Collected')


	def status(self, message):
		print('\r{}'.format(message.ljust(50, ' ')), end = '')
		sys.stdout.flush()

	def progress(self, progress):
		prefix = 'Collecting Grades:'
		size = 50 - len(prefix) - 7
		count = len(self.files_of_interest)
		x = int(size*progress/count)
		print('\r{}[{}{}]{:2d}/{:2d}'.format(prefix, '#' * x, '.' * (size - x), progress, count ), end='')
		sys.stdout.flush()

	def log(self, message):
		if self.development:
			print('{}'.format(message))

DropBoxValueCollector(sys.argv)