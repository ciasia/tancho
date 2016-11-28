import re
import os

def FirstWordFromFilename(file):
	m = re.search("((?:[A-Za-z][A-Za-z]+))", os.path.basename(file))
	return m.group(0)

def GetFileExtension(file):
	# Get file format
	splitFileName = file.split('.')
	return splitFileName[len(splitFileName)-1]