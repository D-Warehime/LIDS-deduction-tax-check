import xlrd
import xlwt
import pymssql
import sys
import csv
import math

conn = pymssql.connect(
	host = 'kaig_prodqa.db.kaiglan.com',
	# port=1438,
	database = 'KAIG',)
cursor = conn.cursor()

companyID = 261931
file_path = '261931Deductions20190711.csv' 

def getSubInfo(employee_id):
	cursor.execute('select subscriber_id from subscriber where company_employee_id = \'{}\' and company_id = \'{}\''.format(employee_id, companyID))
	result = cursor.fetchall()
	return int(result[0][0])

ErrorRecords = []

with open(file_path) as csvfile:
	readCSV = csv.reader(csvfile, delimiter=',')
	for row in readCSV:
		### row format: ['employee id', '', '']
		if len(row) == 3:
			sub_id = getSubInfo(row[0])



		else: 
			ErrorRecords.append(row)



# print(ErrorRecords)

