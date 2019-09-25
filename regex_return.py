#! /Library/Frameworks/Python.framework/Versions/3.4/bin/python3.4
#look through file directory return all file names that include regex
#in the file

import re, os, sys, codecs

def regex_return(*args):
    input_arg = args[0]
    print(input_arg)
    l_args = len(input_arg)
    if l_args == 1:
        print('''Usage:
Provide a directory and a regular expression and file extension.
The tool will parse through the files of directory and return the file names
of any document containing matches to the regular expression.

The expression will return the match as well to help refine expression accuracy. ''')

    if l_args == 2:
        raise RuntimeError('No regular expression provided. Please provide regular expression')

    if l_args == 3:
        raise RuntimeError('Please include file types to parse')
    
    if l_args == 4:
        directory = input_arg[1]
        regex = input_arg[2]
        if not os.path.exists(directory):
            raise RuntimeError('Provided directory does not exist')
            if not os.path.isdir(directory):
                raise RuntimeError('Provided directory is not such.')

        dir_contents = os.listdir(directory)

        ro = re.compile(regex)
        for ffile in dir_contents:
            if os.path.isfile(ffile):
                if ffile.endswith(input_arg[3]):
                    with codecs.open(ffile, 'r', 'utf-8') as f:
                        file_contents = f.read()
                        mo = ro.findall(file_contents)
                        for found in mo:
                            print('File: {}\nContains: {}\n\n'.format(ffile, found))           


        
        
regex_return(sys.argv)




