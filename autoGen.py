from openpyxl import load_workbook, Workbook
import string, itertools

# letters= string.ascii_uppercase
# letters_dic = {i:let for i, let in enumerate(letters, start=0)}
# print letters_dic


wb = load_workbook(filename='data.xlsx')

# grab the active worksheet
sheet1 =wb['Sheet1']


projects=[]

headers= [x.value.encode('UTF8') for x in sheet1.rows[0]]

project_index = headers.index('PROJECT')
transfer_index = headers.index('TRANS_AMT')


headers_dic = {v:k for k,v in enumerate(headers, start=0)}
print headers_dic['PROJECT']

for unit in sheet1.columns[project_index]:
	if unit.value=='PROJECT':
		continue
	if unit.value != None and unit.value not in projects:
		projects.append(unit.value)

projects = [x.encode('UTF8') for x in projects]
projects.sort()
offset= 5


projects_dic = {v:k for (k,v) in enumerate(projects, start=offset)}



new_book = Workbook()
new_sheet = new_book.active
new_sheet.append(projects)
project_col =[]
for col in new_sheet.rows[0]:
	project_col.append(col.column)

project_col_dic = {proj:let for proj,let in zip(projects,project_col)}
#print project_col_dic

for row in sheet1.rows[1:]:
	current_proj=headers_dic['PROJECT']
	current_amt =headers_dic['TRANS_AMT']
	#if row[current_col].value and row[current_col].value != "PROJECT":
	cc = project_col_dic[row[current_proj].value]
	rr = row[current_proj].row
	location = cc+str(rr)
	new_sheet[location]= row[current_amt].value
	

new_book.save('testing.xlsx')
