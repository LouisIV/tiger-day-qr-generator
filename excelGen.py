"""
2018
Louis Lombardo IV
LouisLombardoIV@gmail.com
L_lombarod@u.pacific.edu

Creation of QR codes from Excel documents for:
TIGER CREATIVE, ADMISSIONS
"""

from openpyxl import Workbook
from openpyxl.drawing.image import Image
import qrcode
import csv
import sys

def writeRowToFile(ws, data, current_row, export_location):
	columns = ['A', 'B', 'C', 'D', 'E']
	
	for column in columns:
		ws["%s%s" % (column, current_row)] = data[column]
		
	location = export_location + "/images/" + data['E'] + ".jpeg"
	qr_img = qrcode.make(data['E'])
	qr_img.save(location)

	img = Image(location)
	ws.add_image(img, ('F%s' % current_row))

def create_workbook():
	wb = Workbook()
	ws = wb.active
	return wb, ws

"""
Read CSV file and create .xlsx workbook with the data

csv file
| Term_Descrip | PACIFIC_ID | LAST_NAME | FIRST_NAME | ... | CITY | ... | CURR_1_1_MAJR_DESC | PAC_EMAIL |
|==============|============|===========|============|=====|======|=====|====================|===========|
| Fall 2018    | 9893000000 | Smith     | John       | ... | Apo  | ... | Media X            | --------- |
  ...

.xlsx workbook
|     A     |     B      |   C   |   D  |    E   |  F  |
|===========|============|=======|======|========|=====|
| last_name | first_name | major | city | id_num | img |
  ...

"""
def process_csv(csv_location, export_location, export_filename):
	wb, ws = create_workbook()
	
	with open(csv_location, 'r') as csvfile:
		current_row = 0
		csv_reader = csv.reader(csvfile, delimiter=',', quotechar='|')
		
		for row in csv_reader:
			current_row += 1
			
			if None in (row[2], row[3], row[11], row[7], row[1]):
				raise Exception('One of the values was missing in the CSV file ROW: ', current_row)
			else:
				data = {
					'A': row[2],
					'B': row[3],
					'C': row[15],
					'D': row[7],
					'E': row[1]
				}
				writeRowToFile(ws, data, current_row, export_location)
			
			if current_row > 250:
				break
	
	wb.save(export_location + "/" + export_filename)

if __name__ == "__main__":
	if len(sys.argv) > 3:
		print("CSV Location     : %s" % sys.argv[1])
		print("Export Location  : %s" % sys.argv[2])
		print("Export filename  : %s" % sys.argv[3])
		process_csv(sys.argv[1], sys.argv[2], sys.argv[3])
	else:
		print("Please provide a path to your csv and an export location.")

