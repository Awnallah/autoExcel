from openpyxl import load_workbook, Workbook
import string, itertools

letters= string.ascii_uppercase



wb = load_workbook(filename='data.xlsx')

# grab the active worksheet
sheet1 =wb['Sheet1']


projects=[]

#indeces of headers
headers= [x.value.encode('UTF8') for x in sheet1.rows[0]]

project_index = headers.index('PROJECT')
transfer_index = headers.index('TRANS_AMT')

# make a headers dictionary
headers_dic = {v:k for k,v in enumerate(headers, start=0)}

for unit in sheet1.columns[project_index]:
	if unit.value=='PROJECT':
		continue
	if unit.value != None and unit.value not in projects:
		projects.append(unit.value)

projects = [x.encode('UTF8') for x in projects]
projects.sort()
offset= 5


projects_dic = {v:k for (k,v) in enumerate(projects, start=offset)}

#dictionary for each table header to new excel column letter
first_row_headers = ['DOC_NO','DOC_SUFFIX','FUND_3','POST_DATE']
letters_dic = {let:header_name for let, header_name in zip(first_row_headers,letters)}


first_row = first_row_headers + projects


new_book = Workbook()
new_sheet = new_book.active


new_sheet.append(first_row)
project_col =[]
for col in new_sheet.rows[0]:
	if col.value not in projects:
		continue
	project_col.append(col.column)

project_col_dic = {proj:let for proj,let in zip(projects,project_col)}


for row in sheet1.rows[1:]:
	current_proj = headers_dic['PROJECT']
	current_amt = headers_dic['TRANS_AMT']
	current_doc = headers_dic['DOC_NO']
	current_suffix = headers_dic['DOC_SUFFIX']
	current_fund = headers_dic['FUND_3']
	current_date = headers_dic['POST_DATE']

	# c is for column letter and r is for row number
	c = project_col_dic[row[current_proj].value]
	r = row[current_proj].row
	location = c + str(r)
	new_sheet[letters_dic['DOC_NO'] +str(r)] = row[current_doc].value
	new_sheet[letters_dic['DOC_SUFFIX'] +str(r)] = row[current_suffix].value
	new_sheet[letters_dic['FUND_3'] +str(r)] = row[current_fund].value
	new_sheet[letters_dic['POST_DATE'] +str(r)] = row[current_date].value.date()
	new_sheet[location]= row[current_amt].value

	



new_book.save('testing.xlsx')
