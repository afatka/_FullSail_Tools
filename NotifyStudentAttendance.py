#!/usr/bin/env python3

#!/usr/local/bin/python3
import os, sys, csv, re, smtplib# standard library modules

import openpyxl # Required external module


def read_roster_xlsx( path_to_xlsx_roster):
	''' This function reads the xlsx file and collects student information'''

	roster = openpyxl.load_workbook(path_to_xlsx_roster, data_only = True)

	sheet_names = roster.get_sheet_names()
	sheet_dict = {}

	students_dict = {}
	for row in roster['Attendance']:
		attendance_info_dict = {}

		if row[1].value is not None:
			# print('student number: {}'.format(int(row[1].value)))
			try:
				if int(row[1].value):
					attendance_info_dict['Student Name'] = row[0].value
					attendance_info_dict['Student Number'] = int(row[1].value)
					attendance_info_dict['Total Hours Out'] = int(row[4].value)
					attendance_info_dict['Excused Lecture Hours Out'] = int(row[19].value)
					attendance_info_dict['UnExcused Lecture Hours Out'] = int(row[20].value)
					attendance_info_dict['Lecture Tardies'] = int(row[16].value)

					students_dict[int(row[1].value)] = {'Attendance': attendance_info_dict}
			except ValueError:
				pass

	for row in roster['Lab Attendance']:
		lab_attendance_info_dict = {}

		if row[1].value is not None:
			try:
				if int(row[1].value):
					lab_attendance_info_dict['Student Name'] = row[0].value
					lab_attendance_info_dict['Student Number'] = int(row[1].value)
					lab_attendance_info_dict['Lab Tardy Marks'] = int(row[40].value)
					lab_attendance_info_dict['Excused Lab Hours Out'] = int(row[38].value)
					lab_attendance_info_dict['UnExcused Lab Hours Out'] = int(row[37].value)

					students_dict[int(row[1].value)]['Lab Attendance'] = lab_attendance_info_dict
			except ValueError:
				pass

	#retrieve course code and term ID
	course_code_and_termID = {'Course Code': roster['Attendance']['A2'].value,
								'Term ID': roster['Attendance']['A3'].value}

	# print(students_dict[4735773])
	return students_dict, course_code_and_termID

def read_roster_csv( path_to_csv_roster ):
	''' 
	This funtion reads the csv file and collects 
	student names / email address pairs

	The return dictionary uses student numbers as keys
	'''

	with open(path_to_csv_roster) as csv_file:
		csv_roster = csv.reader(csv_file, delimiter = ',')

		students_dict = {}
		for row in csv_roster:
			student_info_dict ={}

			try: 
				if int(row[0].strip('\'')):
					student_info_dict['Student Number'] = int(row[0].strip('\''))
					student_info_dict['First Name'] = row[1]
					student_info_dict['Last Name'] = row[2]
					student_info_dict['Primary Email'] = row[4]
					student_info_dict['Personal Email'] = row[5]

					students_dict[student_info_dict['Student Number']] = student_info_dict
			except ValueError:
				pass
	return students_dict

def combine_csv_and_xlsx(csv_dict, xlsx_dict):
	'''
	This function takes two dictionaries containing student information
	 and combines them into a single dict.
	CSV is used as master list. MIA's from xlsx are dropped.
	'''
	students_dict = {}
	for student in csv_dict:
		if student in xlsx_dict:
			students_dict[student] = {}
			students_dict[student]['Roster'] = csv_dict[student]
			students_dict[student]['Attendance'] = xlsx_dict[student]['Attendance'] 
			students_dict[student]['Lab Attendance'] = xlsx_dict[student]['Lab Attendance'] 

	# print(students_dict[4735773])
	return students_dict

def construct_emails( text_format, student_info_dict, course_code_and_termID):
	'''
	This function creates the e-mail body for each student
	'''
	with open(text_format, 'r') as f:
		email_format = f.read()

	for student in student_info_dict:
		email_format_edit = email_format
		email_format_edit = re.sub(r'%STUDENT%', '{} {}'.format(
			student_info_dict[student]['Roster']['First Name'], 
			student_info_dict[student]['Roster']['Last Name']),
			email_format_edit)

		email_format_edit = re.sub(r'%COURSECODE%', 
			course_code_and_termID['Course Code'],
			email_format_edit)

		email_format_edit = re.sub(r'%TERMID%', course_code_and_termID['Term ID'], 
			email_format_edit)

		email_format_edit = re.sub(r'%TOTALHOURS%', 
			str(student_info_dict[student]['Attendance']['Total Hours Out']), 
			email_format_edit)

		email_format_edit = re.sub(r'%LECTUREEXUSED%', 
			str(student_info_dict[student]['Attendance']['Excused Lecture Hours Out']), 
			email_format_edit)

		email_format_edit = re.sub(r'%LECTUREUNEXCUSED%', 
			str(student_info_dict[student]['Attendance']['UnExcused Lecture Hours Out']), 
			email_format_edit)

		email_format_edit = re.sub(r'%LABEXCUSED%', 
			str(student_info_dict[student]['Lab Attendance']['Excused Lab Hours Out']), 
			email_format_edit)

		email_format_edit = re.sub(r'%LABUNEXCUSED%', 
			str(student_info_dict[student]['Lab Attendance']['UnExcused Lab Hours Out']), 
			email_format_edit)

		number_of_total_hyphens = 25 - (len(student_info_dict[student]['Roster']['First Name']) +
			 len(student_info_dict[student]['Roster']['Last Name']))

		number_of_excused_spaces = 4 - len(str(student_info_dict[student]['Attendance']['Total Hours Out']))

		email_report = '{0} {1} {4} Total hours: {2} {5} Excused hours: {3}\n'.format(
			student_info_dict[student]['Roster']['First Name'], 
			student_info_dict[student]['Roster']['Last Name'],
			str(student_info_dict[student]['Attendance']['Total Hours Out']),
			str(student_info_dict[student]['Attendance']['Excused Lecture Hours Out'] + 
			student_info_dict[student]['Lab Attendance']['Excused Lab Hours Out']),
			'-' * number_of_total_hyphens, '-' * number_of_excused_spaces
			)

		student_info_dict[student]['Email Info'] = {
			'Email Address': student_info_dict[student]['Roster']['Primary Email'],
			'Message': email_format_edit, 
			'Report': email_report}

	# print(student_info_dict[4735773])
	# print(student_info_dict[4752788])
	return student_info_dict

def send_emails( student_info_dict_emails, course_code_and_termID):
	''' This funciton sends e-mails to the students'''
	try: 

		# THIS IS NOT SECURE!!!
		user = 'folkenfinel@gmail.com'
		pw = 'huqdyjwypdvuchpp' # Requires app password from google

		email_report = '\n\n{0}E-mail Report{0}\n\n'.format('=' * 15)

		server_ssl = smtplib.SMTP_SSL('smtp.gmail.com', 465)
		server_ssl.login(user, pw)


		for student in student_info_dict_emails:
			sent_from = 'afatka@fullsail.com'
			to = student_info_dict_emails[student]['Email Info']['Email Address']
			subject = '{} / {} / - Automated Attendance Update'.format(
				course_code_and_termID['Course Code'], 
				course_code_and_termID['Term ID'])
			message = student_info_dict_emails[student]['Email Info']['Message']

			email_text = 'From: {}\nTo: {}\nSubject: {}\n\n{}'.format(sent_from, to, subject, message)

			server_ssl.sendmail(sent_from, to, email_text)
			del(email_text)
			email_report += student_info_dict_emails[student]['Email Info']['Report']
			print('.', end = '', flush = True)

		return email_report
	except:
		raise

def validate_input_csv_xlsx_txt(*args):
	'''Identify csv file and xlsx file'''
	input_args = args[0]
	if len(input_args) != 3:
		print('Number of files submitted: {}'.format(len(input_args)))
		raise ValueError('Incorrect number of files submitted')

	roster_files = {}

	for f in input_args:
		if f.endswith('.csv'):
			roster_files['csv'] = f
		elif f.endswith('.xlsx'):
			roster_files['xlsx'] = f
		elif f.endswith('.txt'):
			roster_files['txt'] = f
		else:
			raise TypeError('Incorrect file types submitted.')

	return roster_files

def zhu_li():

	print('Starting - Notify Student Attendance')
	print('Validating input', end = '')
	# collect input files for processing
	roster_files = validate_input_csv_xlsx_txt( sys.argv[1:])
	print(' - Completed')
	print('Collecting student e-mail information', end = '')
	# process csv - collect student names and email 
	students_csv_dict = read_roster_csv(roster_files['csv'])
	print(' - Completed')
	print('Collecting student roster information', end = '')
	# process xlsx
	students_xlsx_dict, course_code_and_termID = read_roster_xlsx(roster_files['xlsx'])
	print(' - Completed')
	print('Combining student e-mail and roster information', end = '')
	# combine student info for easier dealing
	student_info_dict_combined = combine_csv_and_xlsx(students_csv_dict, students_xlsx_dict)
	print(' - Completed')
	print('Constructing student e-mails', end = '')
	# construct e-mail
	student_info_dict_emails = construct_emails(roster_files['txt'], 
						student_info_dict_combined, course_code_and_termID)
	print(' - Completed')
	print('Sending student e-mails', end = '', flush = True)
	# send e-mails
	email_report = send_emails(student_info_dict_emails, course_code_and_termID)
	print(' - Completed')
	
	print(email_report)
	print('Task complete.')

zhu_li()