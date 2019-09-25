#!/usr/bin/env python3

# Read three or four colomns from an excel file ( <lastname>, <firstname>, <gradeInt>, <gradeComment> ) 
# Access FSO plateform
# compare student names to those found on portal and post grades

import os, sys, shelve, time
import splinter, xlrd #these are external dependencies

class FSOAuto(object):

	def __init__(self, *args, **kwargs):

		self.development = kwargs.get('development', True)
		# self.lastname_first = kwargs.get('lastnameFirst', True)
		self.login = kwargs.get('login', True)

		self.log('Starting...')
		self.status('Starting...')

		##########################################################
		# All column numbers are INDEX ZERO! #####################
		self.url = kwargs.get('url', 'http://faculty.fso.fullsail.edu')
		self.sheet_of_interest = kwargs.get('sheet', 'sheet1')
		self.starting_row = kwargs.get('starting_row', 2) #skip the top few rows
		self.ignore_cellCol = kwargs.get('ignore_cell', 7)

		self.lastname_CellCol = kwargs.get('lastname_CellCol', 1)
		self.firstname_CellCol = kwargs.get('firstname_CellCol', 2)

		self.gradeInt_CellCol = kwargs.get('gradeInt_CellCol', 3) 
		self.comment_CellCol = kwargs.get('comment_CellCol', 4)
		##########################################################
		

		self.validate_args(args[0]) #validate args and create self.xls_directory variable
		self.student_info = []
		self.collect_student_info()

		if self.login == True:
			self.do_login_info()

		self.launch_browser()

		while True:
			# self.log('Please navigate to page to post grades.')
			i = input('Please navigate to page to post grades.\nCopy the cells you want to post.\nPress Y to continue: ').lower()
			if i == 'y':
				break
		startTime = time.time()
		self.post_grades(self.student_info)
		self.log('Post TIme: {}'.format(round(time.time() - startTime, 2)/60))

		while True:
			i = input('Would you like to post these grades elsewhere? (y/n) ').lower()
			if i == 'y':
				i = input('Please navigate to page to post grades.\nCopy the cells you want to post.\nPress Y to continue: ').lower()
				if i == 'y':
					self.post_grades(self.student_info)
				if i == 'n':
					break
			if i == 'n':
				break

		self.log('break out!')
		self.browser.quit()
		self.log('quit browser')

	def validate_args(self, args):
		self.log('validating args')
		self.log('len: {}'.format(len(args)))
		self.log('args: {}'.format(args))
		if len(args) > 1:
			self.validate_file(args[1])

		if len(args) == 2:
			return 

		if len(args) == 4:
			self.log('updating grade column and comment column')
			self.gradeInt_CellCol, self.comment_CellCol = int(args[2]), int(args[3])
			self.log('grade col: {}'.format(self.gradeInt_CellCol))
			self.log('comment col: {}'.format(self.comment_CellCol))
			return

		raise ValueError('Incorrect Number of inputs provided')

	def validate_file(self, file_to_validate):
		self.log('validate file')
		self.log('file: {}'.format(file_to_validate))
		if not os.path.isfile(file_to_validate):
			raise ValueError('Input is not a file!')
		if not file_to_validate.endswith('.xls'):
			raise ValueError('Wrong file type. \n Supported file types: .xls')
		self.xls_directory = file_to_validate
		return

	def collect_student_info(self):
		wb = xlrd.open_workbook(self.xls_directory)
		for sheet_name in wb.sheet_names():
			self.log('sheet name: {}'.format(sheet_name))
			if sheet_name == self.sheet_of_interest:
				ws = wb.sheet_by_name(sheet_name)
				break
		for row in range(self.starting_row, ws.nrows):
			ignore_row = ws.cell(row, self.ignore_cellCol).value
			if ignore_row == '':
				lastname = ws.cell(row, self.lastname_CellCol).value
				firstname = ws.cell(row, self.firstname_CellCol).value
				grade = ws.cell(row, self.gradeInt_CellCol).value
				comment = ws.cell(row, self.comment_CellCol).value
				if grade != '':
					self.student_info.append(['{}, {}'.format(lastname, firstname), grade, comment])

		for stu in self.student_info:
			self.log('{}:{}, len:{}'.format(stu[0], stu[1], len(stu[2])))

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
		self.browser.visit(self.url)
		if self.login == True:
			self.browser.find_by_id('username').fill(self.username)
			self.browser.find_by_id('password').fill(self.pw)
			self.browser.find_by_text('Login').click()

	def validate_name(self, found_web_name, student_info_list):
		if found_web_name.value == student_info_list[0]:
			return True
		else:
			return False

	def post_grades(self, input_list):
		'''Expected list should be 
		[[student name, grade value, grade comment], 
		[student name, grade value, grade comment]]
		'''
		totalTime = 0
		# is_main = self.browser.is_element_present_by_xpath('//div[@id = "main"]')		
		if self.browser.is_element_present_by_xpath('//div[@id = "batch_grades"]'):
			# self.log('batch?: {}'.format(is_batch_grades))

			student_rows = self.browser.find_by_xpath('//div[@class = "activity row"]')
			# self.log('students found: {}'.format(len(student_rows)))

			iter_count = 1
			for row in student_rows:
				rowStartTime = time.time()
				self.progress(iter_count, len(student_rows))
				iter_count += 1
				student_name = row.find_by_xpath('.//div[@class = "name"]')
				student_grade = row.find_by_xpath('.//input[@id = "grade_score"]')
				student_comment = row.find_by_xpath('.//textarea[@id = "grade_remark"]')

				# self.log('Student: {}'.format(student_name.value))
				# interim_name = student_name.value.split(',')
				# # # self.log('split: {}'.format(interim_name))
				# name_to_match = (interim_name[1] + ' ' + interim_name[0]).strip()
				# self.log('name: {}'.format(repr(name_to_match)))
				for student in input_list:
					# self.log('list name: {}'.format(repr(student[0])))
					if self.validate_name(student_name, student):
					# if name_to_match == student[0]:
						# self.log('NAME MATCH!')
						# self.log('grade: {}'.format(grade))
						student_grade.first.fill(student[1])
						student_comment.first.fill(student[2])
						input_list.remove(student)
				totalTime += round(time.time() - rowStartTime, 2)
		else:
			self.log('batch_grades element not detected.')

		self.log('\naverage post time: {}'.format(round(totalTime/len(student_rows), 2)))
		print('\n')


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

if __name__ == '__main__':
	FSOAuto(sys.argv)
