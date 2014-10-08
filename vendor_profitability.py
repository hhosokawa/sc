from aux_reader import *
import csv
import os
from pprint import pprint
import time

# io
input_dir = 'i/vendor_profitability/'
output = 'o/vendor_profitability.csv'
rows = []

# Pictionary
bi_categories = csv_dic('i/vendor_profitability/auxiliary/bi_categories.csv')
bi_vendors =    csv_dic('i/vendor_profitability/auxiliary/bi_vendors.csv', 2)
categories =    csv_dic('i/vendor_profitability/auxiliary/categories.csv', 2)
divs =          csv_dic('i/vendor_profitability/auxiliary/divs.csv')
departments =   csv_dic('i/vendor_profitability/auxiliary/departments.csv', 2)
gl_parent =     csv_dic('i/vendor_profitability/auxiliary/gl_parent.csv', 2)
vendors =       csv_dic('i/vendor_profitability/auxiliary/vendors.csv')

############### utils ###############

def scan_bi(r):
    r['GL B'] = r['Solution Group']
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

    # Revenue /  FM - GL
    r['GL A'] = 'Field Margin'
    r['Amount'] = float(r['Virtually Adjusted GP'])
    if r['Amount'] != 0:
        rows.append(r.copy())
    r['GL A'] = 'Revenue'
    r['Amount'] = float(r['Virtually Adjusted Imputed Revenue'])
    if r['Amount'] != 0:
        rows.append(r.copy())

def scan_oracle(r, year, qtr):
    r.pop('', None)
    r['Amount'] = r['Amount'].replace(',', '')
    r['Amount'] = float(r['Amount']) * -1
    r['Division'] = divs.get(r['Division'],r['Division'])
    r['GL A'] = gl_parent.get(r['GL Account'], r['Description'])[1]
    r['Quarter'] = qtr
    r['SCC Category'] = categories.get(r['Category'], 'Corporate')[0]
    r['Super Category'] = categories.get(r['Category'], 'Corporate')[1]
    r['Vendor'] = vendors.get(r['Vendor'], r['Vendor'])
    r['Year'] = year

    # Vendor Profitability Description
    dept = r['Department']
    gl = r['GL A']
    if gl in ['SG&A', 'Compensation']:
        if dept == '0922':
            r['GL A'] = 'Total Other Costs'
            r['GL B'] = 'Distribution Services/TDS Costs'
        elif dept == '0403': 
            r['GL A'] = 'Total Other Costs'
            r['GL B'] = 'Licensing Operations'
        elif dept == '0604': 
            r['GL A'] = 'Total Other Costs'
            r['GL B'] = 'Renewals Team'
        elif dept in ['0940', '0941', '0943']:
            r['GL A'] = 'Total Other Costs'
            r['GL B'] = 'PS Resource Costs'
    if dept in departments:
        if gl == 'Compensation':
            r['Super Category'] = departments[dept][1]
            r['GL A'] = 'Cost of Pre Sales & PS Resources'
            r['GL B'] = 'BD Compensation Costs'
        elif gl == 'SG&A':
            r['Super Category'] = departments[dept][1]
            r['GL A'] = 'Cost of Pre Sales & PS Resources'
            r['GL B'] = 'BD SG&A'

    if r['Amount'] != 0:
        rows.append(r)

def scan_csv():
    for file in os.listdir(input_dir):
        file_path = input_dir + file

        # BI - FM
        if file == 'bi.csv':
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
    headers = ['Amount', 'Category', 'Department', 'GL Account', 'GL A', 
               'GL B', 'Description', 'Division', 'Quarter', 'Solution Group',
               'SCC Category', 'Super Category', 'Territory', 'Vendor', 'Year']

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
    print 'vendor_profitability.py complete.', t1-t0
