#!/Applications/Autodesk/maya2017/Maya.app/Contents/bin/mayapy

try: 			
    import maya.standalone 			
    maya.standalone.initialize() 
    print "Maya initialized!"		
except: 			
    raise

import maya.cmds as cmds
import maya.mel as mel
import sys, os, fnmatch


def bundle_and_validate_dirs(args):
	'''This function takes sys.argv, validates input directories and returns them'''

	print('bundle and validate')
	validated_dirs = []
	for arg in args[1:]:
		if os.path.isdir(arg):
			validated_dirs.append(arg)

	return validated_dirs

def collect_maya_files(list_of_dirs):
	'''search recursively through dir and find maya files'''

	print('collect_maya_files')
	file_types_to_find = ['.ma', '.mb']
	maya_files = []
	for folder in list_of_dirs:
		for root, dirnames, filenames in os.walk(folder):
			for ext in file_types_to_find:
				for filename in fnmatch.filter(filenames, '*{}'.format(ext)):
					maya_files.append(os.path.join(root, filename))

	return maya_files

def open_optimize_save(maya_file_path):
	'''take a file, open it in Maya, optimize scene, save as'''
	
	print('open_optimize_save')

	cmds.file(maya_file_path, open = True)

	original_scenename = cmds.file(query = True, sceneName = True)
	print('\n\nSCENE NAME: {}\n\n'.format(original_scenename))

	mel.eval("cleanUpScene 3;")

	cmds.file(rename = os.path.splitext(original_scenename)[0] + '.ma')
	cmds.file(force = True, save = True, type = "mayaAscii")

	cmds.file(force = True, newFile = True)

def process_files():
	''' process files'''

	print('Process Files')

	for file in collect_maya_files(bundle_and_validate_dirs(sys.argv)):
		open_optimize_save(file)

	print('Process Complete')

process_files()




