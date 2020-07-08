import pandas as pd
import numpy as np
import csv
import os

read_file = pd.read_excel (r'Summary_Tables.xlsx', sheet_name='Table 2')
read_file.to_csv (r'Summary_Tables.csv', index = None, header=True)

totals = {
	'Anchorage': [0,0,0,0,0],
	'Gulf Coast': [0,0,0,0,0],
	'Interior':[0,0,0,0,0],
	'Matanuska-Susitna': [0,0,0,0,0],
	'Northern': [0,0,0,0,0],
	'Southeast': [0,0,0,0,0],
	'Southwest': [0,0,0,0,0]
}
region = ''

with open('Summary_Tables.csv') as csvfile:
	readCSV = csv.reader(csvfile, delimiter=',')
	for row in readCSV:
		if row[1] in totals:
			region = row[1]
		if row[3] == 'Total':
			tempTotal = []
			tempTotal.append(int(row[4]))
			tempTotal.append(int(row[11]))
			tempTotal.append(int(row[12]))
			tempTotal.append(int(row[13]))
			tempTotal.append(int(row[10]))
			totals[region] = np.add(totals[region], tempTotal) 
		if row[0] == 'Non-Resident':
			break

columns = ['Region', 'Zip Code', 'Total Cases', 'Deaths', 'Recovered', 'Active', 'Hospitalizations']
zipcodes = [99517, 99669, 99703, 99645, 99762, 99801, 99638]
count = 0
for key in totals:
	totals[key] = np.insert(totals[key], 0, zipcodes[count])
	count+= 1

filename = 'Covid-19 Alaska.csv'
with open(filename, 'w', newline='') as csvfile:
	writer = csv.DictWriter(csvfile, fieldnames = columns)
	writer.writeheader()
	writer = csv.writer(csvfile)
	for key in totals:
		joinedList = [*[key], *totals[key]]
		writer.writerow(joinedList)
		
	
os.remove('Summary_Tables.csv')
os.remove('Summary_Tables.xlsx')
