#from xlrd import open_workbook

import xlrd
book = xlrd.open_workbook('D:\Workspace\Python\ADNIC Dental_Fatima MC_14052015.xlsx')

table = "hello"
sql_base = "INSERT INTO " + table

for sheet in book.sheets():
	sql_base = sql_base + " (" + (', '.join(sheet.row_values(0))) + ") VALUES ("
	
	for row in range(1, sheet.nrows):
		#print(sheet.row_values(row))
		#print(*(sheet.row_values(row)), sep=", ")
		sql_temp = sql_base
		for col in range(sheet.ncols):
			if col != 0:
				sql_temp = sql_temp + ", "

			if sheet.cell(row, col).ctype == xlrd.XL_CELL_TEXT or sheet.cell(row, col).ctype == xlrd.XL_CELL_DATE:
				sql_temp = sql_temp + "'" + sheet.cell(row, col).value + "'"
			elif sheet.cell(row, col).ctype == xlrd.XL_CELL_BOOLEAN or sheet.cell(row, col).ctype == xlrd.XL_CELL_NUMBER:
				sql_temp = sql_temp + str(sheet.cell(row, col).value)
			elif sheet.cell(row, col).ctype == XL_CELL_BLANK or sheet.cell(row, col).ctype == XL_CELL_EMPTY:
				sql_temp = sql_temp + "NULL"
			elif sheet.cell(row, col).ctype == XL_CELL_ERROR:
				print("Corrupted column. Can't insert.")
		sql_temp = sql_temp + ")"
		#print(sql_temp)
		f= open("output.sql","a+")
		f.write(sql_temp+";\n")
			
	

'''
wb = open_workbook('D:\Workspace\Python\Sample.xlsx')
values = []
table = "mytable"
sql = ""
for s in wb.sheets():
    #print 'Sheet:',s.name
    for row in range(1, s.nrows):
        col_names = s.row(0)
        col_value = []
        sql = "INSERT INTO ".append(table).append(" ")
        for name, col in zip(col_names, range(s.ncols)):
            value  = (s.cell(row,col).value)
            try : value = str(int(value))
            except : pass
            col_value.append((name.value, value))
        values.append(col_value)
print(values)
'''