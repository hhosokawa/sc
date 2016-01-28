import csv
import os
from pprint import pprint
import time

# Input / Output
INPUT_DIR = 'i/finance_data/'
OUTPUT = 'P:/_HHOS/DMOO/data/finance_data.csv'

# Converts csv -> Dictionary
def csv_dic(filename, style = 1):     # Converts CSV to dict
    reader = csv.reader(open(filename, "rb"))
    if style == 1: 
        my_dict = dict((k, v) for k, v in reader)
    elif style == 2:
        my_dict = dict((k, (v1, v2)) for k, v1, v2 in reader)
    return my_dict

# Dictionaries
categories = csv_dic('i/finance_data/dictionaries/categories.csv', 2)
divisions = {'100':'Canada', '200':'United States'}
gl_parent = csv_dic('i/finance_data/dictionaries/gl_parent.csv', 2)
territories = csv_dic('i/finance_data/dictionaries/territories.csv')
vendors = csv_dic('i/finance_data/dictionaries/vendors.csv')

############### Data Cleanse Functions ###############
# 1) BI - Create Category A,B,C hierarchy
def clean_bi(r):
    if r['Solution Group'] == 'SERVICES':
        r['Category A'] = r['Solution Type']
    else:
        r['Category A'] = r['Super Category']
        r['Category B'] = r['SCC Category']
        
    # Fix "All Other" Managed Vendor Name for Microsoft and Cisco
    if (r['Super Category'] == 'Microsoft') and (r['Managed Vendor Name'] == 'All Other'):
        r['Managed Vendor Name'] = 'MICROSOFT'
    elif (r['Super Category'] == 'Cisco') and (r['Managed Vendor Name'] == 'All Other'):
        r['Managed Vendor Name'] = 'CISCO SYSTEMS'
        
    # Create USD / Virtually Adjusted PRODUCT / SERVICES GP Columns
    if r['Solution Group'] == 'PRODUCT':
        r['Native Currency Product FM'] = r['Virtually Adjusted GP']
        r['USD Product FM'] = r['USD GP']
    elif r['Solution Group'] == 'SERVICES':
        r['Native Currency Services FM'] = r['Virtually Adjusted GP']
        r['USD Services FM'] = r['USD GP']

    # Assign COGS
    for h in ['USD GP', 'USD Revenue', 'Virtually Adjusted Revenue', 'Virtually Adjusted GP']:
        if '$' in r[h]:
            r[h] = r[h].replace('$', '')
        if ',' in r[h]:
            r[h] = r[h].replace(',', '')
    r['USD COGS'] = float(r['USD Revenue']) - float(r['USD GP'])
    r['Virtually Adjusted COGS'] = float(r['Virtually Adjusted Revenue']) - float(r['Virtually Adjusted GP'])
    
    # Rename GP -> FM
    r['USD FM'] = r['USD GP']
    r['Virtually Adjusted FM'] = r['Virtually Adjusted GP']
    return r
    
# 2) Oracle - Rebate and MDF Revenue / Expense
def clean_oracle(r, book, period, year):
    
    # Clean Formatting
    r.pop('', None)
    if ',' in r['Amount']: 
        r['Amount'] = r['Amount'].replace(',', '')
    r['Book'] = book
    r['Calendar Year'] = year
    r['Division'] = divisions.get(r['Division'], r['Division'])
    r['GL Parent'] = gl_parent.get(r['GL Account'], None)[1]
    r['Category C'] = r['GL Parent']
    r['Fiscal Period'] = period
    r['Managed Vendor Name'] = vendors.get(r['Vendor'], r['Vendor'])
    r['Region'] = territories.get(r['Territory'], r['Territory'])
    # Fix Corporate region -> Canada / US
    if r['Region'] == 'Corporate':
        if r['Division'] == 'Canada':
            r['Region'] = 'Canada Region'
        else:
            r['Region'] = 'US Corporate'
    if r['Category'] in categories:
        r['SCC Category'], r['Super Category'] = categories[r['Category']]
    else:
        r['SCC Category'], r['Super Category'] = r['Category'], r['Category']
    # Assign Quarter
    if 1 <= period <= 3:       r['Fiscal Quarter'] = 1
    elif 4 <= period <= 6:     r['Fiscal Quarter'] = 2
    elif 7 <= period <= 9:     r['Fiscal Quarter'] = 3
    elif 10 <= period <= 12:   r['Fiscal Quarter'] = 4        
        
    # Correct Canada - US East -> Canada - Canada Region
    if (r['Division'] == 'Canada') and (r['Region'] == 'US East Region'):
        r['Region'] = 'Canada Region'
        
    # Fix "All Other" Managed Vendor Name for Microsoft and Cisco
    if (r['Super Category'] == 'Microsoft') and (r['Managed Vendor Name'] == 'All Other'):
        r['Managed Vendor Name'] = 'MICROSOFT'
    elif (r['Super Category'] == 'Cisco') and (r['Managed Vendor Name'] == 'All Other'):
        r['Managed Vendor Name'] = 'CISCO SYSTEMS' 

    # 2013-2014 Cisco Super Category
    if (r['Managed Vendor Name'] == 'CISCO SYSTEMS') and (r['Calendar Year'] in [2013,2014]):
        r['Super Category'] = 'Cisco'
    
    # If Amount != 0, include in rows
    if float(r['Amount']) != 0:
        # Split Rebate, MDF Revenue, Co op Expense into columns
        if book == 'SC Consol - US':
            header_name = 'USD' + ' ' + r['GL Parent']
        else:
            header_name = 'Virtually Adjusted' + ' ' + r['GL Parent']
        r[header_name] = float(r['Amount']) * -1
        
        # Merge Rebate and MDF Revenue -> GP column
        if ('Rebate' in r['GL Parent']) or ('Marketing Revenue' in r['GL Parent']):
            if book == 'SC Consol - US':
                r['USD GP'] = r[header_name]
            else:
                r['Virtually Adjusted GP'] = r[header_name]
        return r

############### Data Output ###############
def scan_csv():
    headers = sorted(['Amount', 'Book', 'Branch', 'Calendar Year', 'Category', 'Category A', 
               'Category B', 'Department', 'Description', 'District', 'Division', 
               'Fiscal Period', 'Fiscal Quarter', 'GL Account', 'GL Parent', 
               'Managed Vendor Name', 'Master ID', 'Master Name', 'Region', 
               'SCC Category', 'Solution Group', 'Solution Type', 'Super Category', 
               'Territory', 'USD FM', 'USD Imputed Revenue', 'USD Revenue', 
               'Unique Master Master Name','Virtually Adjusted FM', 
               'Virtually Adjusted Imputed Revenue', 'Virtually Adjusted Revenue', 
               'Virtually Adjusted Rebates', 'USD Rebates', 'USD Marketing Revenue',
               'Virtually Adjusted Marketing Revenue', 'USD Marketing Expense',
               'Virtually Adjusted Marketing Expense', 'USD COGS', 'Virtually Adjusted COGS',
               'Native Currency Product FM', 'Native Currency Services FM', 
               'USD Product FM', 'USD Services FM', 'OB or TSR', 'USD GP',
               'Virtually Adjusted GP'])

    with open(OUTPUT, 'wb') as o0:
        o0w = csv.DictWriter(o0, delimiter=',',
                             fieldnames=headers, extrasaction='ignore')
        o0w.writerow(dict((fn, fn) for fn in headers))
 
        # Scan *.csv input directory - Data Cleanse
        for file in os.listdir(INPUT_DIR):
            file_path = INPUT_DIR + file

            # 1) BI Sales
            if ('bi' in file) and (file.endswith('.csv')):
                input_file = csv.DictReader(open(file_path))
                for r in input_file:
                    r = clean_bi(r)
                    o0w.writerow(r)
                    
            # 2) Oracle - Rebate and MDF Revenue / Expense
            elif ('SC' in file) and (file.endswith('.csv')):
                book, period, year = file.split('_')
                year = int(year[:-4])    # remove .csv extension
                input_file = csv.DictReader(open(file_path))
                for r in input_file:
                    r = clean_oracle(r, book, period, year)
                    if r: o0w.writerow(r)
    print 'write_csv() complete.'

############### main ###############
if __name__ == '__main__':
    t0 = time.clock()
    scan_csv()
    t1 = time.clock()
    print 'finance_data.py complete.', t1-t0    