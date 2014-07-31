from aux_reader import *
import csv
import os
from pprint import pprint
import time

# io
input_dir = 'i/income_statement_bi/'
output = 'o/income_statement_bi.csv'
rows = []

# Pictionary
bi_categories = csv_dic('i/income_statement_bi/auxiliary/bi_categories.csv')
bi_vendors =    csv_dic('i/income_statement_bi/auxiliary/bi_vendors.csv', 2)
categories =    csv_dic('i/income_statement_bi/auxiliary/categories.csv', 2)
divs =           {'100':'Canada', '200':'United States'}
gl_parent =     csv_dic('i/income_statement_bi/auxiliary/gl_parent.csv', 2)
vendors =       csv_dic('i/income_statement_bi/auxiliary/vendors.csv')

############### utils ###############

def scan_bi(r):
    r['Quarter'] = 'Q' + r['Fiscal Quarter']
    r['Super Category'] = bi_categories.get(r['SCC Category'], r['SCC Category'])
    r['Year'] = r['Calendar Year']

    # BI Vendor -> Oracle Vendor Assignment
    can_us_vendors = ['CISCO PRESS', 'CISCO SYSTEMS', 'EMC CORPORATION', 
                      'HEWLETT PACKARD', 'IBM', 'LENOVO', 'NETAPP']
    if (r['Managed Vendor Name'] in can_us_vendors and
        r['Division'] == 'United States'):
        r['Vendor'] = bi_vendors.get(r['Managed Vendor Name'],
                                     r['Managed Vendor Name'])[1]
    r['Vendor'] = bi_vendors.get(r['Managed Vendor Name'], 
                                 r['Managed Vendor Name'])[0]

    # Revenue / COGS GL
    r['GL Parent'] = 'Net Sales'
    r['Amount'] = float(r['Virtually Adjusted Revenue'])
    if r['Amount'] != 0:
        rows.append(r.copy())
    r['GL Parent'] = 'Cost of Sales'       
    r['Amount'] = -(float(r['Virtually Adjusted Revenue']) - 
                    float(r['Virtually Adjusted GP']))
    if r['Amount'] != 0:
        rows.append(r.copy())

def scan_oracle(r, year, qtr):
    r.pop('', None)
    r['Amount'] = r['Amount'].replace(',', '')
    r['Amount'] = float(r['Amount']) * -1
    r['SCC Category'] = categories.get(r['Category'], 'Corporate')[0]
    div = r['Division']
    r['Division'] = divs.get(div,div)
    r['GL Parent'] = gl_parent.get(r['GL Account'], r['Description'])[1]
    r['Quarter'] = qtr
    r['Super Category'] = categories.get(r['Category'], 'Corporate')[1]
    r['Vendor'] = vendors.get(r['Vendor'], r['Vendor'])
    r['Year'] = year

    if r['Amount'] != 0:
        rows.append(r)

def scan_csv():
    for file in os.listdir(input_dir):
        file_path = input_dir + file

        # BI - FM
        if file == 'bi_fm.csv':
            input_file = csv.DictReader(open(file_path))
            for r in input_file:
                scan_bi(r)
            print file     

        # Oracle - Rebate / MDF
        elif file.endswith(".csv"):
            year, qtr = file.split('_')
            qtr = qtr[:2]
            input_file = csv.DictReader(open(file_path))
            for r in input_file:
                scan_oracle(r, year, qtr)
            print file

def write_csv():
    headers = ['Amount', 'Category',  'GL Account', 'GL Parent', 
               'Description', 'Division', 'Quarter', 'SCC Category', 
               'Super Category', 'Territory', 'Vendor', 'Year']

    with open(output, 'wb') as o0:
        o0w = csv.DictWriter(o0, delimiter=',',
                             fieldnames=headers, extrasaction='ignore')
        o0w.writerow(dict((fn, fn) for fn in headers))
        for r in rows:
            o0w.writerow(r)
    print 'write_csv() completed.'
        
############### main ###############
        
if __name__ == '__main__':
    t0 = time.clock()
    scan_csv()
    write_csv()
    t1 = time.clock()
    print 'income_statement.py complete.', t1-t0
