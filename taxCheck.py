import xlrd
import xlwt
import pymssql
import sys
import csv
import math

conn = pymssql.connect(
	host='kaig_prodqa.db.kaiglan.com',
	# port=1438,
	database='KAIG',)

file_path = '261931Deductions20190711.csv' 

ErrorRecords = []

with open(file_path) as csvfile:
	readCSV = csv.reader(csvfile, delimiter=',')
	for row in readCSV:
		print(row)

