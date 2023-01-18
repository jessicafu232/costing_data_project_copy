# Parsing and data entry
import pandas as pd
from lxml import etree as ET
import argparse
import xlrd
import xlwt
import math
import webbrowser
import subprocess

from datetime import date
import time
import glob
import os

# Notes: There is a ridiculous amount of for nested loops. This could probably be avoided by using numpy vectorization, 
# but the scale of the data is too small for it to really make much of a relevant difference in speed.
# Also, in order to run this program, you need to download your own data file, and move it to the data/ folder. Otherwise, it will automatically use the most
# recent file in the folder (the only demo file currently in the folder is raw_data.xls)

# Finding the latest file in the data folder
list_of_data_files = glob.glob('data/*') 
latest_data_file = max(list_of_data_files, key=os.path.getctime)

# Make it so the user can select which file they would like to read, since the filenames are strange (i.e. 'data/pci_download_1673888103.xls')
parser = argparse.ArgumentParser(description="Collect and store data")
parser.add_argument("--datafile", '-df', default=latest_data_file, help='Location of a raw data file (usually data/myfilename.xls)')
parser.add_argument("--filename", '-fn', default='preview_file', help='The name of the updated data file (without the .xml extension)')
args = parser.parse_args()


# Parsing XML file ahead of time to reference ---------------------------------------------------------------------------------------------------------------------------
tree = ET.parse('costing_data.xml')
root = tree.getroot()

# Checking the most recent year/months in the file ------------------------------------------------------
most_recent_year = root[1].attrib['year']
most_recent_month = root[1].attrib['month']


# Converting data into excel file ---------------------------------------------------------------------------------------------------------------------------------------
# This portion is necessary because the raw data saved from the CEPCI website isn't actually an excel file, it's a text file with a .xls extension (for some reason)
# If CEPCI changes their formatting later, it might affect this part of the code.
book = xlwt.Workbook()
ws = book.add_sheet('Sheet1')  # Add a sheet

f = open(args.datafile, 'r+') # Opening the latest file

data = f.readlines() # Read all lines at once

for i in range(len(data)):
  row = data[i].split('\t')  # Split each column

  for j in range(len(row)):
    ws.write(i, j, row[j])  # Write to cell i, j

book.save('excel_data' + '.xls')
f.close()


# Reading excel file -----------------------------------------------------------------------------------------------------------------------------------------------------

# The data begins at position [2,1], with 2 being the index of the row, and 1 being the index of the column.
costing_dataframe = pd.read_excel(io='excel_data.xls')

# This is the dictionary that will be inputted into the xml file.
# Format is: {'year':{'month': [data, data, data, etc....], 'month': [...]}, 'year':{...}}
input_dict = {}

# Check if the cell is empty.
def isNan(string):
	return string != string

# Initializing variables.
my_month = 'DefaultMonth'
monthly_data = []
my_year = str(int(costing_dataframe.iloc[0,0]))

# Creating a nested dictionary for the year.
input_dict[my_year] = {}

# Running through the file until it reaches the very end.
# The try:except part is just in case it isn't able to read the last value as NaN, it will break out of the loop anyways.

# The data begins at position [2,1].
which_row = 2

while isNan(my_month) is False:
	try:
		# Getting the month of the data
		my_month = costing_dataframe.iloc[which_row,1]

		# Checking that data isn't 0.0 for all values - sometimes by selecting december of this year, even if that data doesn't exist,
		# will create a month that is just zeroes (since the data doesn't exist yet). If that happens, the program will end before that data is entered.
		if costing_dataframe.iloc[which_row, 3] == 0.0:
			break

		# Appending all of the month's data into a list
		for x in range(12):
			monthly_data.append(costing_dataframe.iloc[which_row + x, 3])

		# Appending 
		input_dict[my_year][my_month] = monthly_data
		monthly_data = []

		# Check if the year changes
		if my_month == 'December':
			my_year = str(int(costing_dataframe.iloc[12 + which_row, 0]))
			input_dict[my_year] = {}
			which_row += 1

		which_row += 13

	except: # When it reaches the end of the document and can't read my_month
		break


# Writing into xml file --------------------------------------------------------------------------------------------------------------------------------------------------

# Checking if data overlaps -----------------------------------------------------------------
values_to_delete = []

# This checks whether the most recent month in the costing_data.xml file is in the input_dict (AKA whether they overlap or not)
try:
	input_dict[most_recent_year][most_recent_month]
except: # If not, no problem!
	pass
else: # If so, it deletes the values
	for year in input_dict:
		for month in input_dict[year]:
			values_to_delete.append([year, month])
			if input_dict[year][month] == input_dict[most_recent_year][most_recent_month]:
				break

# Deleting repeated months
for val in values_to_delete:
	del input_dict[val[0]][val[1]]

# Deleting empty years
empty_years = {k: v for k, v in input_dict.items() if not v}
for k in empty_years:
	del input_dict[k]


# Writing to XML --------------------------------------------------------------------------------------------------------------------------------------------------------

# All of the tags we're entering based on the CEPCI format.
tag_list = ['CEIndex', 'Equipment', 'HTXR', 'ProcessMachinery', 'Pipes', 'ProcessInstr', 'Pumps', 'ElectricalEquip', 'Structural', 'ConstructionLabor', 'Buildings', 'EngSupervision']
tag_dict = {}

for year in input_dict:
	for month in input_dict[year]:
		list_of_data = input_dict[year][month]

		# Variables pulled from xls file of raw costing data.
		my_year = year
		my_month = month

		# Making a dictionary with the format of {'CEIndex': '201.6', 'Equipment': '1017.9',...}
		for idx, tag in enumerate(tag_list):
			tag_dict[tag] = str(list_of_data[idx])

		# Creating a new set for the month's data.
		new_set = ET.Element('CostingSet')
		root.insert(1, new_set)
		new_set.set('year', my_year)
		new_set.set('month', my_month)

		# Inserting data.
		for obj in tag_dict:
			new_obj = ET.SubElement(new_set, obj)
			new_obj.text = tag_dict[obj]

			# Indentation / formatting, so the subelements are not all on one line.
			ET.indent(new_set, '	')

			# Inserting a new line.
			new_set.tail = '\n'

# Writing into the new, updated file
tree.write(args.filename+'.xml', pretty_print=True, encoding='utf-8', xml_declaration=True)

print("Wrote into", args.filename + ".xml")


# Checking with the user if they would like to keep this file -----------------------------------------------------------------------------------------------------------
print("Previewing file...")

p = subprocess.Popen(["notepad.exe", args.filename + '.xml'])

time.sleep(1)

user_response = input("Would you like to use this file as the new costing_data.xml? (Respond y / n)\n")

p.kill()

if user_response == 'y':	
	os.rename('costing_data.xml', 'old_costing_data.xml')
	os.rename(args.filename + '.xml', 'costing_data.xml')
	print("Saved updated data into costing_data.xml.")

elif user_response == 'n': 
	print("Saving the file as " + args.filename + ".xml. The old costing_data.xml file will not be overriden.")

print("Saved! Closing program...")
quit()