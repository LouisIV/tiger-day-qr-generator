"""
2018
Louis Lombardo IV
LouisLombardoIV@gmail.com
L_lombarod@u.pacific.edu

Creation of QR codes from csv documents for:
TIGER CREATIVE, ADMISSIONS
"""
import qrcode
import csv
import sys




def create_qr(data, path, title):
	# https://pillow.readthedocs.io/en/3.1.x/reference/Image.html
	img = qrcode.make(data)
	img.save(path + "/" + title)

def process_csv(csv_location, export_location, start_row=False, end_row=False):
	with open(csv_location, 'r') as csvfile:

		# Students.CSV
		# +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		# | Term_Descrip | PACIFIC_ID | LAST_NAME | FIRST_NAME | STREET_LINE1 | ... | PAC_EMAIL |
		# |--------------+------------+-----------+------------+--------------+-----+-----------|
		# | Fall 2018    | 9893000000 | Smith     | John       | N / A        | ... | N / A     |
		#                  |            |           |
		#				   |            |           + > Second part of file name.
		#				   |            + > First part of file name.
		#                  + > Data for the QR generator. 
		# Smith_John.png { 9893000000 }

		current_row = -1
		csv_reader = csv.reader(csvfile, delimiter=',', quotechar='|')
		for row in csv_reader:

			current_row += 1

			if start_row or end_row:
				if current_row < start_row:
					continue
				if current_row >= end_row:
					break
			else:
				# Create the QR Code and save to the export location
				create_qr(row[1], export_location, row[2] + "_" + row[3] + ".jpeg")

if __name__ == "__main__":
	if len(sys.argv) > 2:
		print("CSV Location     : %s" % sys.argv[1])
		print("Export Location  : %s" % sys.argv[2])

		start_row = False
		if len(sys.argv) > 3:
			print("Start Row: %s" % sys.argv[3])
			start_row = sys.argv[3]

		end_row = False
		if len(sys.argv) > 4:
			print("End Row: %s" % sys.argv[4])
			end_row = sys.argv[4]

		process_csv(sys.argv[1], sys.argv[2], start_row=start_row, end_row=end_row)
	else:
		print("Please provide a path to your csv and an export location.")
