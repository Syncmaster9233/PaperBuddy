from openpyxl import Workbook, load_workbook
import csv
import datetime
import os

#opening master template
wb = load_workbook('MST.xlsx')

#getting raw data for parts list going to machine shop
partcount = 0
partlist = []
with open('C:/Users/' + os.getlogin() + '/Downloads/SelectionList.csv') as csv_file:
	#parsing through rows in csv export
	for row in csv_file:

		#skipping first row
		if partcount > 0:
			#parsing row into strings
			rowArr = row.split(",")
			
			#grabbing only info needed part numbers, name and quantity into impArr
			impArr = []
			blockTrack = 0
			for block in rowArr:
				if blockTrack > 0 and blockTrack < 4:
					impArr.append(block)
				blockTrack+=1

			#appending impArr into partlist
			partlist.append(impArr)
		#tracking rows with parts
		partcount +=1	

#consolidating parts
#getting part numbers for unique parts
uniqpartlist = []
for part in partlist:
	unique = True
	for upart in uniqpartlist:
		if part[0] == upart[0]:
			unique = False
	if unique:
		uniqpartlist.append(part)

#summing all unique parts
#probably can be done in previous loop but too lazy
for part in uniqpartlist:
	sum = 0
	for pt in partlist:
		if part[0] == pt[0]:
			sum += int(pt[2])
	part[2] = sum

#finding required number of sheets
uniqPartCount = len(uniqpartlist)
sheetNum = int(uniqPartCount/4) 
if uniqPartCount%4 > 0:
	sheetNum += 1

#importing settings
name = ''
jobNum = ''
model = ''
part = ''
settingsFile = open('Settings.txt', 'r')
for row in settingsFile:
	setarr = row.split(' ')
	for block in setarr:
		if setarr[0] == 'Name:':
			name = setarr[1].replace('\n', '')
		if setarr[0] == 'Job#:':
			jobNum = setarr[1].replace('\n', '')
		if setarr[0] == 'Model:':
			model = setarr[1].replace('\n', '')
		if setarr[0] == 'Part:':
			part = setarr[1].replace('\n', '')

#getting date
dateTemp = datetime.datetime.now()
date = str(dateTemp.month) + '-' + str(dateTemp.day) + '-' + str(dateTemp.year)

Msheet = wb.active

#input settings and date into master template
Msheet['A2'] = str(date)
Msheet['A7'] = str(jobNum)
Msheet['J7'] = str(model)
Msheet['M7'] = str(part)
Msheet['K11'] = str(name)
Msheet['K18'] = str(name)
Msheet['K25'] = str(name)
Msheet['K32'] = str(name)

#copying base sheets for req sheet count
for x in range(1,sheetNum,1):
	wb.copy_worksheet(Msheet)

#cleaning up sheet numbers and page nums and exporting data to spreed sheet
sheetct = 1
for sheet in wb.worksheets:

	#cleaning up names
	sheet.title = 'Sheet'+str(sheetct)
	#page nums input
	sheet['L2'] = 'Page ' + str(sheetct) + ' of ' + str(sheetNum)

	begin = (sheetct-1)*4

	try:
		sheet['A14'] = uniqpartlist[begin][0]
		sheet['A11'] = uniqpartlist[begin][1] + ' (' + str(uniqpartlist[begin][2]) + ')'
	except:
		print('end of uniqpartlist')
	try:
		sheet['A21'] = uniqpartlist[begin+1][0]
		sheet['A18'] = uniqpartlist[begin+1][1] + ' (' + str(uniqpartlist[begin+1][2]) + ')'
	except:
		print('end of uniqpartlist')
	try:
		sheet['A28'] = uniqpartlist[begin+2][0]
		sheet['A25'] = uniqpartlist[begin+2][1] + ' (' + str(uniqpartlist[begin+2][2]) + ')'
	except:
		print('end of uniqpartlist')
	try:
		sheet['A35'] = uniqpartlist[begin+3][0]
		sheet['A32'] = uniqpartlist[begin+3][1] + ' (' + str(uniqpartlist[begin+3][2]) + ')'
	except:
		print('end of uniqpartlist')

	sheetct += 1

wb.save('C:/Users/' + os.getlogin() + '/Desktop/' + date + ' ' + jobNum + '.xlsx')
os.remove('C:/Users/' + os.getlogin() + '/Downloads/SelectionList.csv')