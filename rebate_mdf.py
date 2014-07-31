from aux_reader import *
import csv
import os
from pprint import pprint
import time

# io
input_dir = 'i/rebate_mdf/'
output = 'o/rebate_mdf.csv'
rows = []

# Pictionary
ap =            {'A':'Actual', 'B':'Plan'}
bi_categories = csv_dic('i/rebate_mdf/auxiliary/bi_categories.csv')
bi_vendors =    csv_dic('i/rebate_mdf/auxiliary/bi_vendors.csv', 2)
categories =    csv_dic('i/rebate_mdf/auxiliary/categories.csv')
div =           {'100':'Canada', '200':'United States'}
gl_parent =     csv_dic('i/rebate_mdf/auxiliary/gl_parent.csv')
job_numbers =   csv_dic('i/rebate_mdf/auxiliary/job_numbers.csv', 3)
vendors =       csv_dic('i/rebate_mdf/auxiliary/vendors.csv')

############### utils ###############

def scan_bi(r):
    r['Actual / Plan'] = 'Actual'
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
    r['GL Parent'] = 'Revenue'
    r['Amount'] = float(r['Virtually Adjusted Revenue'])
    if r['Amount'] != 0:
        rows.append(r.copy())
    r['GL Parent'] = 'COGS'       
    r['Amount'] = -(float(r['Virtually Adjusted Revenue']) - 
                    float(r['Virtually Adjusted GP']))
    if r['Amount'] != 0:
        rows.append(r.copy())
    
def scan_oracle(r, actual_plan, year, qtr):
    r.pop('', None)
    r['Actual / Plan'] = ap.get(actual_plan, '')
    r['Amount'] = r['Amount'].replace(',', '')
    r['Amount'] = float(r['Amount']) * -1
    r['Division'] = div.get(r['Division'], r['Division'])
    r['GL Parent'] = gl_parent.get(r['Description'], r['Description'])
    r['Quarter'] = qtr
    r['Super Category'] = categories.get(r['Category'], r['Category'])
    r['Vendor'] = vendors.get(r['Vendor'], r['Vendor'])
    r['Year'] = year

    # Job Number Assignment
    if (r['Project'][:3] in job_numbers or r['Project'] in job_numbers):
        r['Corporate or Custom'] = job_numbers[r['Project']][0]
        r['Marketing Category'] = job_numbers[r['Project']][1]
        r['Marketing Sub Category'] = job_numbers[r['Project']][2]
    else:
        r['Corporate or Custom'] = 'Corporate'
        r['Marketing Category'] = 'Corporate'
        r['Marketing Sub Category'] = 'Corporate'
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
            actual_plan, year, qtr = file.split('_')
            qtr = qtr[:2]
            input_file = csv.DictReader(open(file_path))
            for r in input_file:
                scan_oracle(r, actual_plan, year, qtr)
            print file

def write_csv():
    headers = ['Actual / Plan', 'Amount', 'Corporate or Custom',
               'Description', 'Division', 'GL Parent', 'Marketing Category',
               'Marketing Sub Category', 'Project', 'Quarter', 'Super Category', 
               'Vendor', 'Year']

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
    print 'rebate_mdf.py complete.', t1-t0
