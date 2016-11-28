import sys
import os
from subprocess import call

import helpers
from pricelist import PriceList

originalWorkingDirectory = os.path.dirname(os.path.abspath(sys.argv[1]))

inputFile = sys.argv[1]

mainFileFormat = helpers.GetFileExtension(inputFile)

# Single page csv (Manufacturer is the vendor)
if (mainFileFormat == 'csv'):
	with PriceList(inputFile) as x:
		x.parse()
		x.write()

# Excel file
elif (mainFileFormat[:3] == 'xls'):	

	# Scrape vendor name from filename
	vendorName = helpers.FirstWordFromFilename(inputFile)

	# Create csv files for 
	call(["cscript",os.path.dirname(os.path.abspath(sys.argv[0])) + '\\' + "XlsImport.vbs",os.path.abspath(inputFile)])

	workingDir = os.path.dirname(inputFile)

	fileList = os.listdir(workingDir)

	for file in fileList:
		subFileFormat = helpers.GetFileExtension(file)

		if (subFileFormat == 'csv'):
			with PriceList(originalWorkingDirectory + '\\' + file, ven = vendorName) as x:
				print("Parsing " + workingDir + '\\' + file)
				x.parse()
				x.write()