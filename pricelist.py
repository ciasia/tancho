import xlrd
import csv
import re
import os
import sys

import helpers

class PriceList:

	# Static list of field identifiers/aliases
	wordList = {
		# Can have just a model and no part #, but a part # without a model is just the model
		'Model': ['Product ID', 'Model', 'Part Number'],
		'Part Number': ['Part Number'],
		'Short Description': ['Description'],
		'URL': ['URL'],
		# RRP and Cost are determined by highest and lowest dollar values in sheet
		'MSRP':['RRP', 'MSRP'],
		'Unit Cost':['Unit Cost','Trade','Buy','W/Sale']
	}

	optional_fields = [
		'Part Number',
		'Short Description',
		'URL'
	]



	def __init__(self, file, ven=None):

		# Manufacturer scraped from file name
		self.manufacturer = helpers.FirstWordFromFilename(file)

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

		# Create a list of columns for each desired field
		field_cols = {}
		for key in PriceList.wordList:
			field_cols[key] = -1

		# Increments used for tracking cell_value position
		r = 0
		c = 0

		all_fields_found = False
		all_money_found = False

		header_row = 0

		for row in self.csv_reader:

			highest = float('inf')
			lowest = -1


			if not all_money_found:
				for cell_value in row:
					for field, aliases in PriceList.wordList.items():
						
						# Check cell values against word list
						match = False
						if len(cell_value) > 1: # Only check non-blank cells

							if cell_value[0] == '$':
								castable_cell_value = re.sub('[$,]','',cell_value)
								money = float(castable_cell_value)

								# Find field column by dollar value
								if field == 'rrp':
									if money > highest:
											highest = money
											match = True

								if field == 'cost':
										if money < lowest:
											lowest = money
											match = True				
							else:

								# Find field column by field name
								for alias in aliases:
									reg='.*?('+alias+')'
									m = re.search(reg, cell_value,re.IGNORECASE) #Partial string matching

									if m:
										match = True
										header_row = r

						# Update field columns
						if match:
								field_cols[field] = c

					if not all_fields_found:
						# Check if all required columns have a match
						all_fields_found = True
						for x in field_cols:
							if field_cols[x] == -1 and not (x in PriceList.optional_fields):
								all_fields_found = False
					
					all_money_found = (highest < float('inf') and lowest > -1)

					# Next cell
					c = (c+1)%len(row)

			if all_fields_found and not header_row == r: 

				# Write to internal data
				valid_row = True

				row_data = {}
				row_data['Manufacturer'] = self.manufacturer
				row_data['Vendor'] = self.vendor

				for field in field_cols:
					col = field_cols[field]

					# Detect bad rows
					if row[col] == '' and not (field in PriceList.optional_fields):
						valid_row = False
						break

					# If cell not blank, store value
					if col > -1:
						row_data[field] = row[col]
					else:
						row_data[field] = ''

				if valid_row:
					self.data.append(row_data)

			# Next row
			r += 1

	def write(self):

		### Output directory - to be set by config file in future ###
		if not os.path.basename(os.getcwd()) == 'out':
			if not os.path.exists('out'):
				os.makedirs('out')
			os.chdir('out')

		# Generate file name based on manufacturer/vendor
		newFile = os.getcwd() + '\\' + self.manufacturer + '-' +  self.vendor + ".csv"

		# Create new csv file
		with open(newFile, 'w') as f:
			field_names = ['Manufacturer', 'Vendor'] 
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