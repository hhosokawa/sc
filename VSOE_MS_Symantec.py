import os
import xlrd
import collections
from fuzzywuzzy import fuzz
import time
import csv

output = 'o/Dec 2012 VSOE.csv'
path = 'i/vsoe/'
dirlist = os.listdir(path)
textmatch = fuzz.token_set_ratio

#################################################################################
## Class Declaration

def tree(): return collections.defaultdict(tree)

bundlinfo = collections.namedtuple(
'bundlinfo', 'bundlsku, bundlprice, saname, sasku, saprice, percent')

#################################################################################
## Function

ms_row = 1
symantec_row = 9

# Finders headers column from XLSX Files
def headerfinder(sheet):
	headerdict = {}
	if vendor == 'Microsoft': row = 0
	elif vendor == 'Symantec': row = 8
	for col in range(sheet.ncols):
		cell = sheet.cell_value(row, col)

		# Microsoft Titles
		if cell == 'Product Type': headerdict['Product Type'] = col
		elif cell == 'Product Family': headerdict['Product Family'] = col
		elif cell == 'Pool': headerdict['Pool'] = col
		elif cell == 'Purchase Period': headerdict['Purchase Period'] = col
		elif cell == 'Purchase Unit': headerdict['Purchase Unit'] = col
		elif cell == 'Current Net Price': headerdict['Current Net Price'] = col
		elif cell == 'Level': headerdict['Level'] = col
		elif cell == 'Offering': headerdict['Offering'] = col
		elif cell == 'Program': headerdict['Program'] = col
		elif cell == 'Item Mfg SKU': headerdict['Item Mfg SKU'] = col
		elif cell == 'Item Name': headerdict['Item Name'] = col

		# Symantec Titles
		elif '(From Grouped Tree Price List)' in cell: 
			headerdict['PRODUCT GROUP'] = col
		elif cell == 'MAINTENANCE': headerdict['MAINTENANCE'] = col
		elif cell == 'LICENSE TYPE': headerdict['LICENSE TYPE'] = col
		elif cell == 'PRODUCT FAMILY': headerdict['PRODUCT FAMILY'] = col
		elif cell == 'PROGRAM TYPE': headerdict['PROGRAM TYPE'] = col
		elif cell == 'SHORT DESCRIPTION': headerdict['SHORT DESCRIPTION'] = col
		elif cell == 'MSRP': headerdict['MSRP'] = col
		elif cell == 'PLATFORM': headerdict['PLATFORM'] = col
		elif cell == 'BAND/QTY USERS': headerdict['BAND/QTY USERS'] = col
		elif cell == 'VERSION': headerdict['VERSION'] = col
		elif cell == 'SKU': headerdict['SKU'] = col
		elif cell == 'ITEM TYPE': headerdict['ITEM TYPE'] = col
	return headerdict
						
# Scans XLSX files -> collects to dictionary
def scan(sheet, headerdict, vendor):
	if vendor == 'Microsoft':
		for row in range(ms_row, sheet.nrows):
			revrec = sheet.cell_value(row, headerdict['Product Type'])
			item = ' '.join((str(sheet.cell_value(row, headerdict['Product Family'])), 
  				    str(sheet.cell_value(row, headerdict['Pool'])),
 				    str(sheet.cell_value(row, headerdict['Purchase Period'])),
					str(sheet.cell_value(row, headerdict['Purchase Unit'])),
					str(sheet.cell_value(row, headerdict['Level'])),
					str(sheet.cell_value(row, headerdict['Offering'])),
					str(sheet.cell_value(row, headerdict['Program']))))
			name = sheet.cell_value(row, headerdict['Item Name'])
			price = float(sheet.cell_value(row, headerdict['Current Net Price']))
			try: sku = sheet.cell_value(row, headerdict['Item Mfg SKU'])
			except: sku = sheet.cell_value(row, 4)
			if revrec == 'License/Software Assurance Pack':
				master[item]['SWBNDL'][name] = (price,  sku)
			elif revrec == 'Software Assurance':
				master[item]['SWMTN'][name] = (price, sku)
			elif revrec == 'Standard':
				master[item]['SWLIC'][name] = (price, sku)
	elif vendor == 'Symantec':
		for row in range(symantec_row, sheet.nrows):
			revrec = (sheet.cell_value(row, headerdict['MAINTENANCE']),
					  sheet.cell_value(row, headerdict['LICENSE TYPE']),
					  sheet.cell_value(row, headerdict['ITEM TYPE']))
			item = ' '.join((str(sheet.cell_value(row, headerdict['PRODUCT GROUP'])), 
  				    str(sheet.cell_value(row, headerdict['PRODUCT FAMILY'])),
  				    str(sheet.cell_value(row, headerdict['PROGRAM TYPE'])),
 				    str(sheet.cell_value(row, headerdict['PLATFORM'])),
					str(sheet.cell_value(row, headerdict['BAND/QTY USERS'])),
					str(sheet.cell_value(row, headerdict['VERSION']))))
			name = sheet.cell_value(row, headerdict['SHORT DESCRIPTION'])
			price = float(sheet.cell_value(row, headerdict['MSRP']))
			sku = sheet.cell_value(row, headerdict['SKU'])
			maintenance = sheet.cell_value(row, headerdict['MAINTENANCE'])
			if ('+' in maintenance) and ('LICENSE' in maintenance):
				name = name + ' ' + ' '.join(maintenance.split()[2:])
			else:
				name = name + maintenance
			if ('LICENSE + ' in revrec[0]) and (revrec[1] == 'FULL LICENSE'):
				master[item]['SWBNDL'][name] = (price,  sku)
			elif ('MAINT' in revrec[2]) and (revrec[1] == 'NO LICENSE'):
				master[item]['SWMTN'][name] = (price, sku)
			elif revrec[0] == 'LICENSE ONLY':
				master[item]['SWLIC'][name] = (price, sku)

# Compare SWBNDL to SWMTN items
def bundlmatch(bundldict, sadict):
	for bundl in bundldict:
		for sa in sadict:
			
			# If Match HAS NOT been found
			if bundldict[bundl].__class__ == tuple: 
				bundldict[bundl] = bundlinfo( 
								   bundlsku = bundldict[bundl][1],
								   bundlprice = bundldict[bundl][0],
								   saname = sa, 
								   sasku = sadict[sa][1],
								   saprice = sadict[sa][0], 
								   percent = sadict[sa][0]/bundldict[bundl][0])

			# If match HAS been found
			else:
				old_bundlsku,old_bundlprice,old_saname,old_sasku,_,_ = bundldict[bundl]
				if textmatch(bundl, sa) > textmatch(bundl, old_saname):
					bundldict[bundl] = bundlinfo(
									   bundlsku = old_bundlsku,
                                       bundlprice = old_bundlprice,
									   sasku = old_sasku,
									   saname = sa, 
								   	   saprice = sadict[sa][0], 
								   	   percent = sadict[sa][0]/old_bundlprice)
	return bundldict

#################################################################################
## Main

master = tree()

# Phase 1: Scrape all excel data
t0 = time.clock()
for f in dirlist:
	try:
		if 'NAM_' in f: vendor = 'Symantec'
		else: vendor = 'Microsoft'
		book = xlrd.open_workbook(''.join([path, f]))
		sheet = book.sheet_by_index(0)
		headerdict = headerfinder(sheet)
 		scan(sheet, headerdict, vendor)
	except IOError:
		print 'Failed opening: ', f
		continue
print 'Phase 1: Data Scrape Complete'

# Phase 2: Parse correct matches between Bundles - Stand Alone items
for item in master:
	mibndl = master[item]['SWBNDL']
	milic = master[item]['SWLIC']
	mimtn = master[item]['SWMTN']
	if mibndl and mimtn:
		master[item]['SWBNDL'] = bundlmatch(mibndl, mimtn)
print 'Phase 2: Data Analysis Complete'
		
# Phase 3: Output Correct Parses
with open(output, 'wb') as o1:
	wr = csv.writer(o1)
	wr.writerows([('Bundle Name', 'Bundle SKU', 'Bundle $', 
				'SA Name', 'SA SKU', 'SA $', 'Percent')])
	for item in master:
		for bundl in master[item]['SWBNDL']:
			if master[item]['SWBNDL'][bundl].__class__ == tuple:
				pass
			else:
				i = master[item]['SWBNDL'][bundl]
				wr.writerow([bundl, i.bundlsku, i.bundlprice, 
							i.saname, i.sasku, i.saprice, 
							i.percent])
print 'Phase 3: Writing to Disk... Complete'
t1 = time.clock()
print 'Process completed! Duration:', t1-t0


		
