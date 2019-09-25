#!/usr/bin/env python3

# Read three or four colomns from an excel file ( <lastname>, <firstname>, <gradeInt>, <gradeComment> ) 
# Access FSO plateform
# compare student names to those found on portal and post grades

import os, sys, shelve, time
import splinter, xlrd #these are external dependencies

# phantomjs driver required for headless operation download from http://phantomjs.org/

class FSOAuto(object):

	def __init__(self, *args, **kwargs):

		self.development = kwargs.get('development', True)
		self.login = kwargs.get('login', True)
		#self.browser_type = 'firefox' #'phantomjs'
		self.browser_type = 'phantomjs'

		self.safe_submit_stall = True
		self.quit_on_end = True

		self.log('Starting...')
		self.status('Starting AutoPost')

		##########################################################
		# All column numbers are INDEX ZERO! #####################
		self.url = kwargs.get('url', 'http://faculty.fso.fullsail.edu')
		self.sheet_of_interest = kwargs.get('sheet', 'roster')
		self.starting_row = kwargs.get('starting_row', 1) #skip the top few rows
		self.course_code_cellCol = kwargs.get('course_code_cellCol', 1)
		self.section_cellCol = kwargs.get('section_cellCol', 2) #cell to watch for section number
		self.ignore_cellCol = kwargs.get('ignore_cell', 17) #column to seek for ignore text

		self.term_id_Cell = kwargs.get('term_id', (1, 1)) # location of the term ID on the xls (row, column) ##INDEX ZERO!!!!
		self.project_name_row = kwargs.get('project_row', 3) #name of the project / Should be above the grade value(Int)

		self.lastname_CellCol = kwargs.get('lastname_CellCol', 1)
		self.firstname_CellCol = kwargs.get('firstname_CellCol', 2)

		self.gradeInt_CellCol = kwargs.get('gradeInt_CellCol', 3) #column of grade value (Int)
		self.comment_CellCol = kwargs.get('comment_CellCol', 4) #column of grade comment (str)
		##########################################################
		if self.gradeInt_CellCol == self.course_code_cellCol or self.gradeInt_CellCol == self.section_cellCol:
			raise ValueError('Course Code and Section Code cell CANNOT be the same as gradeInt Cell')
		
		self.status('Validating Inputs')
		self.validate_args(args[0]) #validate args and create self.xls_directory variable
		self.sections_info = []
		self.status('Collecting Worksheet Info')
		self.collect_worksheet()
		self.get_project_name()
		self.collect_student_info()

		# self.log('Diagnostics')
		# for section_to_post in self.sections_info:
		# 	self.log('{}:{}'.format(section_to_post[0], section_to_post[1]))
		# 	self.log('Project: {}\n'.format(self.project_name))
		# 	for student in section_to_post[2]:
		# 		self.log('{}'.format(student[0]))
		# 	self.log('\n')
		# raise ValueError('Completed print')

		if self.login == True:
			self.do_login_info()

		self.status('Launching Browser')
		self.launch_browser()
		for i in range(4):
			self.log('sleep...Zzz')
			time.sleep(1)

		self.log('number of sections found: {}'.format(len(self.sections_info)))
		t = '.'
		if len(self.sections_info) > 1:
			t = 's.'
		self.status('Handling {} Section{}'.format(len(self.sections_info), t))
		for section_to_post in self.sections_info:
			self.status('Posting {}:{}\n'.format(section_to_post[0], section_to_post[1]))
			self.nav_from_home_to_section(section_to_post[0], section_to_post[1])
			self.log('made it to section?: {}:{}'.format(section_to_post[0], section_to_post[1]))
			for i in range(4):
				self.log('sleep...Zzz')
				time.sleep(1)
			self.nav_from_section_to_project()
			self.post_grades(section_to_post[2])

			if self.browser_type == 'firefox' and self.safe_submit_stall:
				while True:
					i = input('Press "y" to continue and submit.').lower()
					if i == 'y':
						break

			self.log('attempting submit')
			self.you_must_submit()
			self.log('did it submit?')

			time.sleep(1)
			self.nav_to_home()

		self.log('finished posting')

		self.log('do quit.')
		self.status('Finished Posting\n')
		if self.quit_on_end:
			self.browser.quit()

	def you_must_submit(self):
		self.log('Submit to me now!')
		self.browser.find_by_xpath('//a[@id = "batch_save_btn"]')[0].click()
		time.sleep(5)
		self.browser.find_by_xpath('//div[@class = "modal-footer"]//a[@class = "btn btn large btn-primary"]')[0].click()

	def find_classes(self, *args):
		self.log('find classes')
		if self.browser.is_element_present_by_xpath('//div[@id = "courses"]'):
			self.browser.find_link_by_text('All').click()
			courses = lambda: self.browser.find_by_xpath('//div[@id = "courses"]')

			self.validate_term_id(courses().find_by_xpath('./header//div[@class = "filter-title"]').value)
			
			course_headers = courses().find_by_xpath('//div[@class = "course-header grid-12"]')
			course_sections = courses().find_by_xpath('//ul[@class = "course-grid"]')
			zipped_list = list(zip(course_headers, course_sections))
			self.courses_w_sections = []
			for c in zipped_list:
				course_code = c[0].find_by_xpath('./h2').value # course header div element
				sections = c[1].find_by_xpath('.//div[@class = "course-section-name"]/a') # course sections ul element

				self.courses_w_sections.append([course_code, sections])

	def nav_to_term_id(self, term_id_input, *args):
		self.log('incorrect term id')
		while True:
			i = input('Please navigate to term {}\nPress Y to continue'.format(term_id_input)).lower()
			if i == 'y':
				break
		time.sleep(1)
		self.find_classes()

	def nav_from_section_to_project(self):
		time.sleep(1)
		self.log('navigate to project / batch grade')
		self.log('\n\nfind batch grade!')
		self.log('click project link')
		link = self.browser.find_by_text('{}'.format(self.project_name))
		link = '{}{}'.format(link['href'], '/batch_grade')
		self.log('href: {}'.format(link))
		self.browser.visit(link)

	def nav_from_home_to_section(self, courseCode, sectionID):
		self.log('Navigate from home to section')
		self.log('Course: {}'.format(courseCode))
		self.log('Section: {}'.format(sectionID))
		self.log('starting url: {}'.format(self.browser.url))
		
		self.find_classes()

		for course in self.courses_w_sections:
			self.log('testing course[0]: {}'.format(course[0]))
			if course[0] == courseCode:
				self.log('course match: {}'.format(courseCode))
				for section in course[1]:
					self.log('testing section: {}'.format(section.value))
					if section.value == sectionID:
						self.log('section match: {}'.format(sectionID))
						tempUrl = self.browser.url
						section.click()
						self.log('did the click...did it go?')
						time.sleep(1)
						if self.browser.url == tempUrl:
							section.click()
						break
				break
		assert self.browser.url.endswith('activities')

	def nav_to_home(self):
		self.log('Navigate Home')
		time.sleep(3)
		self.browser.visit(self.url)
		i = 0
		self.log('entering while loop')
		while i < 3:
			time.sleep(5)
			self.log('url: {}'.format(self.browser.url))
			if self.browser.url.endswith(self.url[7:] + '/'):
				self.log('url correct, break')
				break
				self.log('{}: {}'.format(i,self.browser.url))
			self.log('{}: url not correct, go to {}'.format(i,self.url))
			self.browser.visit(self.url)
			i += 1

		self.log('did I go home?')
		self.log('url: {}'.format(self.browser.url))
		self.log('home: {}'.format(self.url))
		self.log('slice: {}'.format(self.url[7:]) + '/')
		assert self.browser.url.endswith(self.url[7:] + '/')

	def validate_term_id(self, term_id_input, *args):
		term_id_fromXL = str(int(self.ws.cell(self.term_id_Cell[0], self.term_id_Cell[1]).value))
		term_id_fromWeb = term_id_input

		if term_id_fromWeb.endswith(term_id_fromXL):
			return
		self.nav_to_term_id(term_id_fromXL)

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

	def get_project_name(self):
		self.log('get project name')
		self.project_name = self.ws.cell(self.project_name_row, self.gradeInt_CellCol).value
		if self.project_name == None or self.project_name == '':
			raise ValueError('No Project Name Found!\n{}'.format(self.project_name))
		self.log('Project: {}'.format(self.project_name))
		while True:
			i = input('\nProject is: {}\nIs this correct? (y/n)'.format(self.project_name)).lower()
			if i == 'y':
				break 
			if i == 'n':
				raise ValueError('Incrrect Project')

	def collect_worksheet(self):
		self.log('collect_worksheet')
		wb = xlrd.open_workbook(self.xls_directory)
		self.ws = None
		for sheet_name in wb.sheet_names():
			self.log('sheet name: {}'.format(sheet_name))
			if sheet_name == self.sheet_of_interest:
				self.ws = wb.sheet_by_name(sheet_name)
				break
		if self.ws == None:
			raise ValueError('No sheet named {}'.format(self.sheet_of_interest))

	def collect_student_info(self):
		self.log('collect student info')
		section = []
		student_info =[]
		for row in range(self.starting_row, self.ws.nrows):
			ignore_row = self.ws.cell(row, self.ignore_cellCol).value
			# self.log('ignore row: {}'.format(repr(ignore_row)))
			if ignore_row == '':
				# self.log('ignore row != ""')
				if any(i.isdigit() for i in self.ws.cell(row, self.course_code_cellCol).value):
					# self.log('course code found')
					if section != []:
						section.append(student_info)
						self.sections_info.append(section)
						section = []
						student_info = []

					course_code = self.ws.cell(row, self.course_code_cellCol).value
					section_value = self.ws.cell(row, self.section_cellCol).value
					# self.log('{} - {}'.format(course_code, section_value))
					section.append(course_code)
					section.append(section_value)
				else:
					lastname = self.ws.cell(row, self.lastname_CellCol).value
					firstname = self.ws.cell(row, self.firstname_CellCol).value
					grade = self.ws.cell(row, self.gradeInt_CellCol).value
					comment = self.ws.cell(row, self.comment_CellCol).value
					if grade != '' and ignore_row == '':
						student_info.append(['{}, {}'.format(lastname, firstname), grade, comment])
		section.append(student_info)
		self.sections_info.append(section)

	def do_login_info(self):
		directory = os.path.dirname(__file__)
		db = os.path.join(directory, 'autoPostAuth')

		if os.path.isfile(db + '.db'):
			with shelve.open(db, flag = 'r') as s:
				self.username = s['username']
				self.pw = s['password']
		else:
			while True:
				i = input('Would you like me to remember your username and password? (y/n): ').lower()
				if i == 'y':
					self.username = input('What is your username? ')
					self.pw = input('What is your password? ')
					with shelve.open(db) as s:
						s['username'] = self.username
						s['password'] = self.pw
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
		if self.browser_type == 'firefox':
			self.browser = splinter.Browser() #firefox
		if self.browser_type == 'phantomjs':
			self.browser = splinter.Browser('phantomjs')#phantomjs
			self.browser.driver.set_window_size(1400, 800)

		self.browser.visit(self.url)
		if self.login == True:
			self.browser.find_by_id('username').fill(self.username)
			self.browser.find_by_id('password').fill(self.pw)
			self.browser.find_by_text('Login').click()
		else:
			while True:
				i = input('Please login to continue. \nPress Y to continue:').lower()
				if i == 'y':
					break

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
		# totalTime = 0
		if self.browser.is_element_present_by_xpath('//div[@id = "batch_grades"]'):
			student_rows = self.browser.find_by_xpath('//div[@class = "activity row"]')

			iter_count = 1
			for row in student_rows:
				# rowStartTime = time.time()
				# self.progress(iter_count, len(student_rows))
				# iter_count += 1
				student_name = row.find_by_xpath('.//div[@class = "name"]')
				student_grade = row.find_by_xpath('.//input[@id = "grade_score"]')
				student_comment = row.find_by_xpath('.//textarea[@id = "grade_remark"]')

				for student in input_list:
					if self.validate_name(student_name, student):
						self.progress(iter_count, len(input_list) + iter_count - 1)
						iter_count += 1
						student_grade.first.fill(student[1])
						student_comment.first.fill(student[2])
						input_list.remove(student)
				# totalTime += round(time.time() - rowStartTime, 2)
		else:
			self.log('batch_grades element not detected.')

		print('\n')

	def status(self, message):
		if not self.development:
			print('\r{}'.format(message.ljust(50, ' ')), end = '')
			sys.stdout.flush()
		else:
			print('{}'.format(message))

	def progress(self, progress, length):
		prefix = 'Processing Students:'
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
