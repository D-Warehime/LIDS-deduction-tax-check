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

def getSubInfo(employee_id):
	cursor.execute("select subscriber_id, state from subscriber where company_employee_id = \'{}\' and company_id = \'{}\'".format(employee_id, companyID))
	result = cursor.fetchall()
	if not int(result[0][0]):
		return 0, ''
	state = result[0][1]
	#Using a sales tax rate of 8% for Ontario province and 9% for Quebec, these are for the Medical and Dental plans
	if state == 'ON':
		TaxRate = 0.08
	elif state == 'QC':
		TaxRate = 0.09
	else: 
		TaxRate = 0
	return int(result[0][0]), TaxRate

companyID = 261931
# Dfile_path = '261931Deductions20190711.csv' 
deduction_file_path = 'Dtest.csv'
benefit_file_path = 'Btest.csv'

benefit_file = {}
with open(benefit_file_path) as bfile:
	benefitFile = csv.reader(bfile, delimiter=',')
	for row in benefitFile:
		if row[0] != 'Employee_number':
			benefit_sub_id, x = getSubInfo(int(row[0]))
			if benefit_sub_id not in benefit_file:
				benefit_file[benefit_sub_id] = []
			benefit_file[benefit_sub_id].append({row[1]: float(row[2])})

ErrorRecords = []
Non_QC_ON_Records = []
QC_ON_TaxRateRecords = []
rowLine = 0

with open(deduction_file_path) as csvfile:
	readCSV = csv.reader(csvfile, delimiter=',')
	currentSub = {"sub": 0, "plan_and_rates": {}}
	for row in readCSV:
		if rowLine > 0:
		### row format: [employee id, company plan type, rate]
			if len(row) == 3:
				sub_id, TaxRate = getSubInfo(row[0])

				if sub_id == 0:
					ErrorRecords.append(row)
					continue
				if TaxRate == 0:
					Non_QC_ON_Records.append(row)
					continue
				if currentSub["sub"] != sub_id:
					currentSub["sub"] = [sub_id]
					currentSub["plan_and_rates"] = {}

				if row[1] not in currentSub["plan_and_rates"]:
					currentSub["plan_and_rates"][row[1]] = row[2]
				else: 
					ErrorRecords.append(row)
					continue

				if 'H Sales Tax' in currentSub["plan_and_rates"] and 'EE HEALTH' in currentSub["plan_and_rates"]:
					amount = TaxRate * currentSub["plan_and_rates"]['EE HEALTH']
					if amount == currentSub["plan_and_rates"]['H Sales Tax']:
						QC_ON_TaxRateRecords.append(row)
						continue
					else:
						ErrorRecords.append(row)
						continue

				if 'D Sales Tax' in currentSub["plan_and_rates"] and 'EE DENTAL' in currentSub["plan_and_rates"]:
					amount = TaxRate * currentSub["plan_and_rates"]['EE DENTAL']
					if amount == currentSub["plan_and_rates"]['D Sales Tax']:
						QC_ON_TaxRateRecords.append(row)
						continue
					else:
						ErrorRecords.append(row)
						continue

				if 'H ER Sales Tax' in currentSub["plan_and_rates"]:
					if sub_id in benefit_file:
						if 'ER HEALTH' in benefit_file[sub_id]:
							amount = TaxRate * benefit_file[sub_id]['ER HEALTH']
							if amount == row[3]:
								QC_ON_TaxRateRecords.append(row)
					else:
						ErrorRecords.append(row)

				if 'D ER Sales Tax' in currentSub["plan_and_rates"]:
					if sub_id in benefit_file:
						if 'ER DENTAL' in benefit_file[sub_id]:
							amount = TaxRate * benefit_file[sub_id]['ER DENTAL']
							if amount == row[3]:
								QC_ON_TaxRateRecords.append(row)
					else:
						ErrorRecords.append(row)
		else: 
			if rowLine != 0:
				ErrorRecords.append(row)
		rowLine += 1

print("Error records", len(ErrorRecords), ErrorRecords)
print("Sales Tax records", len(QC_ON_TaxRateRecords), QC_ON_TaxRateRecords)
print("Not tax records", len(Non_QC_ON_Records), Non_QC_ON_Records)

#Save results of this script into an excel sheet
resultBook = xlwt.Workbook()
sheet1 = resultBook.add_sheet('Error records')
sheet2 = resultBook.add_sheet('QC & ON records')
sheet3 = resultBook.add_sheet('Non QC & ON records')
for row, record in enumerate(ErrorRecords):
	for column, value in enumerate(record):
		sheet1.write(row, column, value)
for row, record in enumerate(QC_ON_TaxRateRecords):
	for column, value in enumerate(record):
		sheet2.write(row, column, value)
for row, record in enumerate(Non_QC_ON_Records):
	for column, value in enumerate(record):
		sheet3.write(row, column, value)
resultBook.save('testResults.xls')