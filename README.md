
# Tiger Day QR Generator

These python scripts will create QR codes from csv documents.

- [generator.py](generator.py) will create these qr codes in a specified output folder
- [excelGen.py](excelGen.py) will create a .xlsx workbook like this:

|     A     |     B      |   C   |   D  |    E   |  F  |
| --------- | ---------- |------ | ---- | ------ | --- |
| last_name | first_name | major | city | id_num | img |

But these values can be anything you want from the csv file. You can edit which column maps to which like this:
```python
# `row` represents the entire row from the csv
# `row[0]` = the first column of the current row
data = {
	'A': row[2],
	'B': row[3],
	'C': row[15],
	'D': row[7],
	'E': row[1]
}
writeRowToFile(worksheet, data, current row, export location)
```
> Note that you also need to edit which field is converted to a qr code in `writeRowToFile`. Default is `E`

[excelGen.py](excelGen.py) will also create a folder containing all of the QR codes in the same directory as the .xlsx workbook file.
