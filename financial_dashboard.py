from aux_reader import *
from datetime import datetime
from dateutil.parser import parse
from pprint import pprint
import os
import time

# io
input_dir = 'i/financial_dashboard/'
output = 'o/financial_dashboard.csv'
rows = []

# Pictionary
categories = csv_dic('i/financial_dashboard/auxiliary/categories.csv', 2)
divisions = {'100':'Canada', '200':'United States'}
gl_parent = csv_dic('i/rebate_mdf/auxiliary/gl_parent.csv', 2)
territories = csv_dic('i/financial_dashboard/auxiliary/territories.csv')
vendors = csv_dic('i/financial_dashboard/auxiliary/vendors.csv')

############### Data Cleanse ###############
# Scan Input Directory
def scan_csv():
    for file in os.listdir(input_dir):
        file_path = input_dir + file

        # BI - Field Margin
        if file == 'bi.csv':
            input_file = csv.DictReader(open(file_path))
            for r in input_file:
                rows.append(r)
                
        # Oracle - Rebate and MDF Revenue / Expense
        elif file.endswith(".csv"):
            year, period = file.split('_')
            period = int(period[:-4])    # remove .csv extension
            input_file = csv.DictReader(open(file_path))
            for r in input_file:
                clean_oracle(r, year, period)
    print 'scan_csv() complete.'

# Oracle - Rebate and MDF Revenue / Expense
def clean_oracle(r, year, period):
    
    # Clean Formatting
    r.pop('', None)
    if ',' in r['Amount']: 
        r['Amount'] = r['Amount'].replace(',', '')
    r['Calendar Year'] = year
    r['Division'] = divisions.get(r['Division'], r['Division'])
    r['GL Parent'] = gl_parent.get(r['GL Account'], None)[1]
    r['Fiscal Period'] = period
    if 1 <= period <= 3:       r['Fiscal Quarter'] = 1
    elif 4 <= period <= 6:     r['Fiscal Quarter'] = 2
    elif 7 <= period <= 9:     r['Fiscal Quarter'] = 3
    elif 10 <= period <= 12:   r['Fiscal Quarter'] = 4
    r['Managed Vendor Name'] = vendors.get(r['Vendor'], r['Vendor'])
    r['Region'] = territories.get(r['Territory'], r['Territory'])
    if r['Region'] == 'Corporate':
        if r['Division'] == 'Canada':
            r['Region'] = 'Canada Region'
        else:
            r['Region'] = 'US Corporate'
    if r['Category'] in categories:
        r['SCC Category'], r['Super Category'] = categories[r['Category']]
    else:
        r['SCC Category'], r['Super Category'] = r['Category'], r['Category']

    # If Amount != 0, include in rows
    if float(r['Amount']) != 0:
        if r['GL Parent'] == 'Rebates':
            r['USD Rebate'] = float(r['Amount']) * -1
        elif ((r['GL Parent'] == 'Marketing Revenue') or 
            (r['GL Parent'] == 'Marketing Expense')):
            r['USD MDF GP'] = float(r['Amount']) * -1
        rows.append(r)

############### Data Output ###############
def write_csv():
    headers = ['Calendar Year', 'Category', 'Department', 'Description', 'Division', 
               'Fiscal Period', 'Fiscal Quarter', 'GL Account', 'Managed Vendor Name', 
               'Region', 'SCC Category', 'Solution Group', 'Solution Type', 
               'Super Category', 'USD GP', 'USD Imputed Revenue', 'USD MDF GP', 
               'USD Rebate', 'USD Revenue']

    with open(output, 'wb') as o0:
        o0w = csv.DictWriter(o0, delimiter=',',
                             fieldnames=headers, extrasaction='ignore')
        o0w.writerow(dict((fn, fn) for fn in headers))
        for r in rows:
            o0w.writerow(r)
    print 'write_csv() complete.'
             
############### main ###############
if __name__ == '__main__':
    t0 = time.clock()
    scan_csv()
    write_csv()
    t1 = time.clock()
    print 'sales_expense_report.py complete.', t1-t0