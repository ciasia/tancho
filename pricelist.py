import xlrd
import csv
import re
import os
import sys

import helpers

class PriceList:


	# Static list of field identifiers/aliases
	wordList = {
		'Model': ['Product ID', 'Model', 'Part Number', 'Part'],
		'Part Number': ['Part Number'],
		'Short Description': ['Description'],
		'URL': ['URL'],
		'MSRP':['RRP', 'MSRP'],
		'Unit Cost':['Unit Cost','Trade','Buy','W/Sale']
	}

	# RRP and Cost are determined by highest and lowest dollar values in sheet
	# Archival values for possible future use
	'''
	'MSRP':['RRP', 'MSRP'],
	'Unit Cost':['Unit Cost','Trade','Buy','W/Sale']
	'''

	optional_fields = [
		'Part Number',
		'Short Description',
		'URL',
		'MSRP'
	]



	def __init__(self, file, ven=None):

		print("Parsing " + file)

		# Manufacturer scraped from file name
		splitFile = file.split('\\')
		fileName = splitFile[len(splitFile) -1]
		self.manufacturer = fileName.split('.')[0]

		# Vendor passed by method caller
		self.vendor = ven or self.manufacturer

		# Input file path
		self.csv_path = os.path.abspath(file)
		
		# Excel file
		self.csv_file = open(self.csv_path, 'r', errors='ignore', encoding='utf-8')
		
		# Reader for file
		self.csv_reader = csv.reader(self.csv_file)	

		# Internal representation of data
		self.data = []

	def parse(self):

		# Create a list(map) of columns for each desired field
		# Initially each desired field has a column value of -1 (null)
		field_cols = {}
		for key in PriceList.wordList:
			field_cols[key] = -1

		# Increments used for tracking current cell position in source worksheet
		r = 0
		c = 0

		# Whether the columns associated to all desired data fields have been identified
		all_fields_found = False

		# Row containing column names (assumed to be zero until proven otherwise)
		header_row = 0

		for row in self.csv_reader:

			highest = -1 
			lowest = float('inf')

			if not all_fields_found:
				# Look for desired field columns
				for cell_value in row:
					# Only check non-blank cells
					if len(cell_value) > 1: 

						for field, aliases in PriceList.wordList.items():
							
							match = False
							currency_regex = re.compile("^\s*\$*\s*[0-9]+,*\s*[0-9]*\.+[0-9]+\s*$")

							# Find field column by field name
							if not currency_regex.match(cell_value):
								for alias in aliases:
									if field_cols[field] == -1:
										reg='.*?('+alias+')'
										m = re.search(reg, cell_value,re.IGNORECASE) #Partial string matching

										if m:
											match = True
											header_row = r

							# Find field column by dollar value		
							else:
								castable_cell_value = re.sub("[$,\s]",'',cell_value)
								money = float(castable_cell_value)

								
								if field == 'MSRP':
									if money > highest:
											highest = money
											match = True

								if field == 'Unit Cost':
										if money < lowest:
											lowest = money
											match = True

							# Update field columns
							if match:
									field_cols[field] = c

						# Check if all required columns have a match
						if not all_fields_found:
							all_fields_found = True
							for x in field_cols:
								if field_cols[x] == -1 and not (x in PriceList.optional_fields):
									all_fields_found = False

					# Next cell
					c = (c+1)%len(row)


			# Check if columns containing data have already been found
			if all_fields_found and not header_row == r: 

				# Assume row contains desired values until proven otherwise
				valid_row = True

				# Temporary dictionary for desired values of current row
				row_data = {}
				row_data['Manufacturer'] = self.manufacturer
				row_data['Vendor'] = self.vendor

				for field, field_col in field_cols.items():

					# Detect bad rows
					if row[field_col] == '' and not (field in PriceList.optional_fields):
						valid_row = False
						break

					# If cell not blank, grab value
					if field_col > -1:
						row_data[field] = row[field_col]
					else:
						row_data[field] = ''

				# If row contains valid data, store in self.data
				if valid_row:

					# "Can have just a Model and no Part Number, but a Part Number without a model is just the model"
					if 'Part Number' in row_data:
						if not row_data['Part Number'] == '':

							if row_data['Model'] == '':
								row_data['Model'] = row_data['Part Number']
								row_data['Part Number'] = ''

							elif row_data['Model'] == row_data['Part Number']:
								row_data['Part Number'] = ''
					
					if 'MSRP' in row_data and 'Unit Cost' in row_data:
						if row_data['MSRP'] == row_data['Unit Cost']:
							row_data['MSRP'] = ''

					self.data.append(row_data)


			# Next row
			r += 1

		#print(field_cols)

		for field, col in field_cols.items():
			if not field in self.optional_fields and field_cols[field] == -1:
				print("Missing data: "+field)

		# Return false if no data found
		return ((len(self.data) > 0))

	def write(self):

		### Output directory - to be set by config file in future version ###
		if not os.path.basename(os.getcwd()) == 'out':
			if not os.path.exists('out'):
				os.makedirs('out')
			os.chdir('out')

		# Generate file name based on manufacturer/vendor
		newFile = os.getcwd() + '\\' + self.vendor + '-' +  self.manufacturer + ".csv"

		# Create new csv file
		with open(newFile, 'w') as f:
			field_names = ['Vendor', 'Manufacturer'] 
			field_names.extend(PriceList.wordList.keys())
			w = csv.DictWriter(f, fieldnames=field_names, lineterminator='\n')
			w.writeheader()
			w.writerows(self.data)

	# Make class compatible with Python's 'with' statement
	def __enter__(self):
		return self

	def __exit__(self, exc_type, exc_value, traceback):
		self.csv_file.close()
		os.unlink(self.csv_path)