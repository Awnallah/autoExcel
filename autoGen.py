from openpyxl import load_workbook, Workbook
import string, itertools

# letters= string.ascii_uppercase
# letters_dic = {i:let for i, let in enumerate(letters, start=0)}
# print letters_dic


wb = load_workbook(filename='data.xlsx')

# grab the active worksheet
sheet1 =wb['Sheet1']

# for column in sheet_ranges.rows:
# 	print column[0].value

arr= list(dir(sheet1.columns))
#print arr
projects=[]

headers= [x.value.encode('UTF8') for x in sheet1.rows[0]]

project_index = headers.index('PROJECT')
transfer_index = headers.index('TRANS_AMT')
print transfer_index

headers_dic = {v:k for k,v in enumerate(headers, start=0)}
print headers_dic['PROJECT']
print headers_dic['TRANS_AMT']

for unit in sheet1.columns[project_index]:
	if unit.value=='PROJECT':
		continue
	if unit.value != None and unit.value not in projects:
		projects.append(unit.value)

projects = [x.encode('UTF8') for x in projects]
projects.sort()
offset= 5
#print (projects)
# arls=["1","2","3","4"]
#projects_dic = {projects.index(k):k for k in (projects)}
projects_dic = {v:k for (k,v) in enumerate(projects, start=offset)}



new_book = Workbook()
new_sheet = new_book.active
new_sheet.append(projects)
project_col =[]
for col in new_sheet.rows[0]:
	project_col.append(col.column)

project_col_dic = {p:l for p,l in zip(projects,project_col)}
print project_col_dic

# for row in sheet1.rows:
# 	if row[headers_dic['PROJECT']].value and row[headers_dic['PROJECT']].value != "PROJECT":
# 		cc = projects_dic[row[headers_dic['PROJECT']].value]
# 		print cc
	# for cell in row:
	# 	if cell.value != None:
	# 		r=cell.row;c=cell.column
	# 		cr= str(c)+str(r)
	# 		new_sheet[cr]=cell.value

new_book.save('testing.xlsx')
