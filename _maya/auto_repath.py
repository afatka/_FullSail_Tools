import maya.cmds as cmds
import maya.mel as mel
import os
def auto_repath():
	version = 'Beta 0.1'
	repath_paths = []
	# unresolved_plugs = mel.eval('filePathEditor -query -unresolved -listFiles "" -withAttribute;')[1::2] 

	file_path = cmds.file(query = True, sceneName = True)
	file_split = os.path.split(file_path)
	repath_paths.append(file_split[0])
	# print(file_split[0])

	max_depth = 6
	starting_depth = 0
	for directoryName, subDirectoryName, fileList in os.walk(file_split[0]):
		# print('dir name: {}'.format(directoryName))
		if starting_depth == 0:
			starting_depth = len(directoryName.split(os.sep))
		directory_depth = len(directoryName.split(os.sep))
		if directory_depth - starting_depth >= max_depth:
			print('os.walk recursion break')
			break

		for subDir in subDirectoryName:
			if 'normal' in subDir.lower():
				repath_paths.insert(0, os.path.join(directoryName, subDir))
	for i in repath_paths:
		# print(i)
		unresolved_plugs = mel.eval('filePathEditor -query -unresolved -listFiles "" -withAttribute;')[1::2] 
		# print(' '.join(unresolved_plugs))
		mel.eval('filePathEditor -repath "{}" -recursive {}'.format(i, ' '.join(unresolved_plugs)))

auto_repath()