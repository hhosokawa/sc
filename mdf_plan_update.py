from aux_reader import *
from datetime import datetime
from dateutil.parser import parse
from pprint import pprint
import os
import time

# io
input_dir = 'i/mdf_plan_update/'
output = 'o/mdf_plan_update.csv'
rows = []

# Pictionary
ap =  {'A':'Actual', 'B':'Plan'}
categories = csv_dic('i/mdf_plan_update/auxiliary/categories.csv', 3)
divisions = {'100':'Canada', '200':'United States'}
gl_parent = csv_dic('i/mdf_plan_update/auxiliary/gl_parent.csv', 2)
vendors = csv_dic('i/mdf_plan_update/auxiliary/vendors.csv')

############### Data Cleanse ###############
# Scan Input Directory
def scan_csv():
    for file in os.listdir(input_dir):
        file_path = input_dir + file

        # Oracle - MDF Revenue / Expense
        if file.endswith(".csv"):
            year, period, _, _, actual_plan = file.split('_')
            period = int(period)
            actual_plan = actual_plan[:-4]
            input_file = csv.DictReader(open(file_path))
            for r in input_file:
                clean_oracle(r, actual_plan, year, period)
    print 'scan_csv() complete.'

# Oracle - Rebate and MDF Revenue / Expense
def clean_oracle(r, actual_plan, year, period):

    if ',' in r['Amount']: 
        r['Amount'] = r['Amount'].replace(',', '')
    
    # If Amount != 0, include in rows, and clean formatting
    if float(r['Amount']) != 0:
        r.pop('', None) 
        r['Actual / Plan'] = ap.get(actual_plan, '')
        r['Amount'] = float(r['Amount']) * -1
        r['Division'] = divisions.get(r['Division'], r['Division'])
        r['GL Parent'] = gl_parent.get(r['GL Account'], None)[1]
        r['Period'] = period
        if 1 <= period <= 3:       r['Quarter'] = 1
        elif 4 <= period <= 6:     r['Quarter'] = 2
        elif 7 <= period <= 9:     r['Quarter'] = 3
        elif 10 <= period <= 12:   r['Quarter'] = 4    
        r['Year'] = year
        r['Vendor Name'] = vendors.get(r['Vendor'], r['Vendor'])
        r = get_category(r)
        rows.append(r)
        
# Oracle - Assign Category
def get_category(r):
    if r['Category'] in categories:
        r['SCC Category'] = categories[r['Category']][1]        
        r['Super Category'] = categories[r['Category']][2]
    else:
        r['SCC Category'], r['Super Category'] = 'Other', 'Other'
    
    # Departments exclude 0701, 0704 -> Other
    if r['Department'] in ['0701', '0704']:
        r['SCC Category'], r['Super Category'] = 'Other', 'Other'
        
    # Data Center Broad Portfolio
    if r['Category'] in ['201','203','204','205','207','492']:
        broad_portfolio_vendors = ['0001', '0002', '0011', '0012', '0030', '0034', 
                                   '0067', '0069', '0070', '0075', '0076', '0089']
        if r['Vendor'] in broad_portfolio_vendors:
            r['SCC Category'] = 'Data Center Broad Portfolio'
            r['Super Category'] = 'DC & Hybrid IT'

    # Category Cisco <2014 using Vendor field
    if float(r['Year']) < 2015:
        if r['Vendor'] in ['0006', '0007']:
            r['Super Category'] = 'DC & Hybrid IT'
            r['SCC Category'] = 'Cisco'
    return r
    
############### Data Output ###############
def write_csv():
    headers = ['Actual / Plan','Amount', 'Category', 'Department', 
               'Description', 'Division', 'GL Account', 'GL Parent', 
               'Period', 'Project', 'Quarter', 'SCC Category', 
               'Super Category', 'Territory', 'Vendor', 
               'Vendor Name', 'Year']

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
