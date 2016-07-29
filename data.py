
import csv, os, codecs
#os.makedirs('updated', exit_ok=True)
csvRows=[]
csvFileObj = open('trial.csv', 'rb')
readerObj = csv.DictReader(csvFileObj)

for row in readerObj:
	if row['PROJECT'] not in csvRows:
		csvRows.append(row['PROJECT'])

csvRow = csvRows.sort()
print csvRows


output= open('output.csv','wb')
writer = csv.writer(output)
writer.writerow("")
writer.writerow(csvRows)


