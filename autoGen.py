from openpyxl import load_workbook, Workbook
wb = load_workbook(filename='data.xlsx')

# grab the active worksheet
sheet1 =wb['Sheet1']

# for column in sheet_ranges.rows:
# 	print column[0].value

arr= list(dir(sheet1.columns))
#print arr
projects=[]
headers= list([x.value for x in sheet1.rows[0]])

project_index = headers.index('PROJECT')

for unit in sheet1.columns[project_index]:
	if unit.value=='PROJECT':
		continue
	if unit.value != None and unit.value not in projects:
		projects.append(unit.value)

projects = [x.encode('UTF8') for x in projects]
projects.sort()
#print (projects)
# arls=["1","2","3","4"]
#projects_dic = {projects.index(k):k for k in (projects)}
projects_dic = {k:v for (k,v) in enumerate(projects, start=5)}
#print projects_dic


new_book = Workbook()
new_sheet = new_book.active
new_sheet.append(projects)
#print (new_sheet['D5'].col_idx)
new_book.save('testing.xlsx')
