import openpyxl
#this package makes it easy to work with excel files

#read the excel file row by row and checks if the zip code is default then it simply adds the count of it
#otherwise it simply puts 0 against the zip if it is not defaulted
#everytime it encounters a same zip code it sees if it is defaulted it adds 1 to its count else ignore
#dat is a dictionary to hold zip codes (as strings) keys and count of defaults as values

workbook = openpyxl.load_workbook('data.xlsx')
sheet = workbook.active

dat = {}
for row in range(2, sheet.max_row + 1):
	d = sheet['F'+str(row)].value
	z = sheet['H'+str(row)].value
	if z != '' or z != ' ':
		if z not in dat:
			dat[z] = d
		else:
			if d == 1:
				dat[z] = dat[z] + 1

#count the maximum number of the defaults accross any zip
total = 0
for e in dat:
	if dat[e]> total:
		total = dat[e]

#probs is a dictionary to hold zip codes (as strings) keys and the probability as values
probs = {}
for e in dat:
    probs[e] = dat[e] / float(total)

#open the sheet2 of our workbook
sh2 = workbook.get_sheet_by_name('Sheet2')
sh2['A1'] = "Zip Code"
sh2['B1'] = "Probability to default"

#write the zip codes and probability to the sheet
count = 2
for e in probs:
    sh2['A' + str(count)] = str(e)
    sh2['B' + str(count)] = str(probs[e])
    count = count + 1
#saves sheet with new name
workbook.save('data2.xlsx')
