import sys
import os
from subprocess import call

import helpers
from pricelist import PriceList

inputFile = sys.argv[1]
inputFileDirectory = os.path.dirname(os.path.abspath(sys.argv[1]))
inputFileFormat = helpers.GetFileExtension(inputFile)

# Single page csv (Manufacturer is the vendor)
if (inputFileFormat == 'csv'):
	with PriceList(inputFile) as x:
		x.parse()
		x.write()

# Excel file
elif (inputFileFormat[:3] == 'xls'):	

	# Scrape vendor name from filename
	vendorName = helpers.FirstWordFromFilename(inputFile)

	if not os.path.exists('temp'):
		os.makedirs('temp')

	# Create csv files for parsing
	call(["cscript", os.path.dirname(os.path.abspath(sys.argv[0])) + '\\' + "XlsImport.vbs", os.path.abspath(inputFile), os.path.abspath('temp')])

	tempDir = os.path.abspath('temp')

	fileList = os.listdir(tempDir)

	for file in fileList:
		subFileFormat = helpers.GetFileExtension(file)

		if (subFileFormat == 'csv'):
			with PriceList(tempDir + '\\' + file, ven = vendorName) as x:
				x.parse()
				x.write()