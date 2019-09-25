#!/usr/bin/env python3

# Read two colomns from an excel file
# Access FSO plateform
# compare student names to those found on portal and post grades

import os, sys, shelve, shlex
import pyperclip, splinter #these are external dependencies

class FSOAuto(object):

	def __init__(self, *args, **kwargs):
		# test_list = [['Emily Aguilar', '84.1', 'Test Comment one'],
		# 			['Jensen Alicea Class', '80', 'another comment'],
		# 			['Bradley Allen', '50.71', 'and a third']]

		self.development = kwargs.get('development', True)
		self.lastname_first = kwargs.get('lastnameFirst', True)
		self.login = kwargs.get('login', True)


		self.log('Starting...')
		self.status('Starting...')

		if self.login == True:
			self.do_login_info()

			# TODO - launch FSO plateform, load to page, find students
		self.launch_browser()

		while True:
			# self.log('Please navigate to page to post grades.')
			i = input('Please navigate to page to post grades.\nCopy the cells you want to post.\nPress Y to continue: ').lower()
			if i == 'y':
				break

		self.convert_clipboard_to_list()
		self.post_grades(self.info_list)

		while True:
			i = input('Would you like to post more grades? (y/n) ').lower()
			if i == 'y':
				i = input('Please navigate to page to post grades.\nCopy the cells you want to post.\nPress Y to continue: ').lower()
				if i == 'y':
					self.convert_clipboard_to_list()
					self.post_grades(self.info_list)
				if i == 'n':
					break
			if i == 'n':
				break

		self.log('break out!')
		self.browser.quit()
		self.log('quit browser')

		# self.convert_clipboard_to_list()
		# for student in self.info_list:
		# 	self.validate_name('student, Joe', student)

	# def convert_clipboard_to_list(self):
	# 	# self.log('convert clipboard')
	# 	clipboard = pyperclip.paste()
	# 	# self.log('clipboard: {}'.format(repr(clipboard)))
	# 	clipboard_list = clipboard.split('\r')
	# 	# self.log('list: {}'.format(clipboard_list))
	# 	self.info_list = []
	# 	for info in clipboard_list:
	# 		temp = info.split('\t')
	# 		if temp[-1].isdigit():
	# 			temp.append('')
	# 		self.info_list.append(temp)
	# 	for student in self.info_list:
	# 		self.log(student)

	def convert_clipboard_to_list(self):
		# self.log('converting to clipboard')
		clipboard = pyperclip.paste()
		# self.log('cb: \n{}'.format(clipboard))
		clip_list = shlex.split(clipboard)
		# self.log('clip_list: {}'.format(clip_list))
		self.info_list =[]
		student_depth = 4
		if clip_list[1].isdigit():
			self.log('one cell for the name')
			student_depth = 3
		i = 1
		temp = []
		for list_item in clip_list:
			temp.append(list_item)
			if i % student_depth == 0:
				self.info_list.append(temp)
				# self.log('temp: {}'.format(temp))
				temp = []
			i += 1

		# for student in self.info_list:
		# 	self.log(student)




	def do_login_info(self):
		self.log('manage login info')
		directory = os.path.dirname(__file__)
		self.log('directory: {}'.format(directory))
		self.log(os.path.join(directory, 'autoPostAuth.db'))
		db = os.path.join(directory, 'autoPostAuth')

		if os.path.isfile(db + '.db'):
			self.log('Authentification file found.')	
			with shelve.open(db, flag = 'r') as s:
				self.username = s['username']
				self.pw = s['password']

			self.log('un: {}'.format(self.username))
			self.log('pw: {}'.format('*'*len(self.pw)))
		else:
			self.log('No File Found')
			while True:
				i = input('Would you like me to remember your username and password? (y/n): ').lower()
				if i == 'y':
					self.username = input('What is your username? ')
					self.pw = input('What is your password? ')
					with shelve.open(db) as s:
						s['username'] = self.username
						s['password'] = self.pw
					self.log('un: {}'.format(self.username))
					self.log('pw: {}'.format('*'*len(self.pw)))
					while True: 
						i2 = input('Does this seem correct? (y/n): ').lower()
						if i2 == 'n':
							i = 'blah'
							break
						if i2 == 'y':
							break
				if i == 'y' or i == 'n':
					break

	def launch_browser(self):
		self.browser = splinter.Browser()
		self.browser.visit('http://faculty.fso.fullsail.edu')
		if self.login == True:
			self.browser.find_by_id('username').fill(self.username)
			self.browser.find_by_id('password').fill(self.pw)
			self.browser.find_by_text('Login').click()

	def validate_name(self, found_web_name, student_info_list):
		name_cells = []
		for i in range(len(student_info_list)):
			if student_info_list[i].isdigit():
				break
			name_cells.append(student_info_list[i])
		self.log('name_cells: {}'.format(name_cells))
		self.log('len: {}'.format(len(name_cells)))

		if len(name_cells) > 2:
			if self.lastname_first:
				firstname = name_cells[-1]
				name_cells.remove(firstname)
				temp = ''
				for name in name_cells:
					temp += name
				lastname = temp
			else:
				firstname = name_cells[0]
				name_cells.remove(firstname)
				temp = ''
				for name in name_cells:
					temp += name
				lastname = temp

		if len(name_cells) == 2:
			firstname, lastname = name_cells
			if self.lastname_first:
				lastname, firstname = name_cells
		if len(name_cells) ==1:
			firstname, lastname = name_cells[0].split('_')
			if self.lastname_first:
				lastname, firstname = name_cells[0].split('_')

		# firstname = name_cells[0] # CDC O first name
		# lastname = name_cells[1] # CDC O last name
		

		copied_name = (firstname + ' ' + lastname).strip()
		self.log('copied_name: {}'.format(copied_name))
		 
		
		interim_name = found_web_name.value.split(',')
		name_to_match = (interim_name[1] + ' ' + interim_name[0]).strip()
		
		self.log('name_to_match: {}'.format(name_to_match))
		self.log('copied_name: {}'.format(copied_name))
		if name_to_match == copied_name:
			return True
		else:
			return False

	def post_grades(self, input_list):
		'''Expected list should be 
		[[student name, grade value, grade comment], 
		[student name, grade value, grade comment]]
		'''
		# is_main = self.browser.is_element_present_by_xpath('//div[@id = "main"]')		
		if self.browser.is_element_present_by_xpath('//div[@id = "batch_grades"]'):
			# self.log('batch?: {}'.format(is_batch_grades))

			student_rows = self.browser.find_by_xpath('//div[@class = "activity row"]')
			# self.log('students found: {}'.format(len(student_rows)))

			iter_count = 1
			for row in student_rows:
				self.progress(iter_count, len(student_rows))
				iter_count += 1
				student_name = row.find_by_xpath('.//div[@class = "name"]')
				student_grade = row.find_by_xpath('.//input[@id = "grade_score"]')
				student_comment = row.find_by_xpath('.//textarea[@id = "grade_remark"]')

				# self.log('Student: {}'.format(student_name.value))
				interim_name = student_name.value.split(',')
				# # self.log('split: {}'.format(interim_name))
				name_to_match = (interim_name[1] + ' ' + interim_name[0]).strip()
				# self.log('name: {}'.format(repr(name_to_match)))
				for student in input_list:
					# self.log('list name: {}'.format(repr(student[0])))
					if self.validate_name(student_name, student):
					# if name_to_match == student[0]:
						# self.log('NAME MATCH!')
						# self.log('grade: {}'.format(grade))
						student_grade.first.fill(student[-2])
						student_comment.first.fill(student[-1])
						input_list.remove(student)
		else:
			self.log('batch_grades element not detected.')

		print('\n')

		# self.browser.quit()
	def status(self, message):
		if not self.development:
			print('\r{}'.format(message.ljust(50, ' ')), end = '')
			sys.stdout.flush()

	def progress(self, progress, length):
		prefix = 'Processing Rows:'
		size = 50 - len(prefix) - 7
		count = length
		x = int(size*progress/count)
		print('\r{}[{}{}]{:2d}/{:2d}'.format(prefix, '#' * x, '.' * (size - x), progress, count ), end='')
		sys.stdout.flush()

	def log(self, message):
		if self.development:
			print('{}'.format(message))


FSOAuto(sys.argv)
