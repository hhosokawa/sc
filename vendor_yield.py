from aux_reader import *
import csv
import os
from pprint import pprint
import time

# io
input_dir = 'i/vendor_yield/'
output = 'o/vendor_yield.csv'
rows = []

# Pictionary
categories = csv_dic('i/vendor_yield/hierarchies/categories.csv')
gl_parent = csv_dic('i/vendor_yield/hierarchies/gl_parent.csv')
vendors = csv_dic('i/vendor_yield/hierarchies/vendors.csv')

# Vendor Lists
top_vendors = ['MICROSOFT', 'CISCO SYSTEMS', 'DELL COMPUTER',
               'VMWARE', 'LENOVO', 'IBM', 'ADOBE SYSTEMS',
               'NETAPP', 'SYMANTEC', 'HP', 'EMC CORPORATION',
               'MCAFEE', 'SOPHOS', 'CITRIX SYSTEMS', 'APPLE',
               'RED HAT SOFTWARE', 'TREND MICRO']
HP_fix = ['HP INC', 'HEWLETT-PACKARD ENTERPRISE', 'HEWLETT PACKARD']

############### utils ###############
# Include / Edit Top Vendors
def vendor_name_fix(managed_vendor_name):
    managed_vendor_name_top = ''
    # Consolidated HP
    if managed_vendor_name in HP_fix:
        managed_vendor_name_top = 'HP'
    # Only include vendors from top_vendors list
    elif managed_vendor_name in top_vendors:
        managed_vendor_name_top = managed_vendor_name
    else:
        managed_vendor_name_top = 'All Other'
    return managed_vendor_name_top 

# BI
def scan_bi(r):
    # Adjust BI Super Categories
    if (r['Super Category'] == 'Professional Services (INACTIVE)' 
        or r['Solution Group'] == 'SERVICES'):
        r['Super Category'] = 'Services'
    if r['SCC Category'] in ['Enterprise Software', 'Security']:
        r['Super Category'] = 'Enterprise SW & Security'
    if (r['Solution Group'] == 'PRODUCT' 
        and r['Solution Type'] == 'SaaS - Cloud'):
        r['Super Category'] = 'SaaS'
    r['Managed Vendor Name Top'] = vendor_name_fix(r['Managed Vendor Name'])
    
    r['Amount'] = r['USD GP']
    r['GL Parent'] = 'FM'
    rows.append(r.copy())
    r['Amount'] = r['USD Imputed Revenue']
    r['GL Parent'] = 'Imputed Revenue'
    rows.append(r.copy())
    r['Amount'] = r['USD Revenue']
    r['GL Parent'] = 'Revenue'
    rows.append(r.copy())
    
# Oracle - SS
def scan_oracle(r, year, qtr):
    r.pop('', None)
    r['Amount'] = r['Amount'].replace(',', '')
    r['Amount'] = float(r['Amount']) * -1
    if r['Amount'] != 0:
        r['Calendar Year'] = int(year)
        r['Fiscal Quarter'] = int(qtr)
        r['GL Parent'] = gl_parent.get(r['GL Account'],'')
        r['Managed Vendor Name'] = vendors.get(r['Vendor'], 'All Other')
        r['Managed Vendor Name Top'] = vendor_name_fix(r['Managed Vendor Name'])
        r['Super Category'] = categories.get(r['Category'],'')
        
        # 2014 Cisco Super Category
        if ('CISCO' in r['Managed Vendor Name'] and r['Calendar Year'] == 2014 
            and r['Department'][:-2] == '07'):
            r['Super Category'] = 'Cisco'
        
        # BD Department Filtering
        if r['GL Parent'] == 'Rebate':
            if r['Department'] in ['0701', '0704']:
                r['Super Category'] = 'Other'
            elif r['Department'][-2:] != '07':
                r['Super Category'] == 'Services'
        rows.append(r)

def scan_csv():
    for file in os.listdir(input_dir):
        file_path = input_dir + file

        # BI
        if 'bi' in file:
            input_file = csv.DictReader(open(file_path))
            for r in input_file:
                scan_bi(r)
            print file
            
        # Oracle - SS
        elif file.endswith(".csv"):
            year, qtr = file.split('_')
            qtr = qtr[:2]
            input_file = csv.DictReader(open(file_path))
            for r in input_file:
                scan_oracle(r, year, qtr)
            print file
            
def create_zeros():
    for quarter in range(1,5):
        r = {}
        r['Amount'] = 0
        r['Calendar Year'] = 2015
        r['Fiscal Quarter'] = quarter
        r['Managed Vendor Name'] = 'All Other'
        r['Super Category'] = 'Other'
        for gl_parent in  ['Imputed Revenue', 'Revenue', 'FM', 'MDF GP', 'Rebate']:
            r['GL Parent'] = gl_parent
            rows.append(r.copy())
    print 'ceate_zeros() complete.'
            
def write_csv():
    headers = ['Amount', 'Category', 'Calendar Year', 'Department', 
               'Description', 'Division', 'Fiscal Quarter', 'GL Account', 
               'GL Parent', 'Managed Vendor Name', 'Managed Vendor Name Top',
               'SCC Category', 'Solution Group', 'Solution Type', 
               'Super Category', 'Territory']
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
    create_zeros()
    write_csv()
    t1 = time.clock()
    print 'vendor_yield.py complete.', t1-t0