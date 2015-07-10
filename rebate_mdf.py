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
categories =    csv_dic('i/rebate_mdf/auxiliary/categories.csv', 3)
departments =   csv_dic('i/rebate_mdf/auxiliary/departments.csv', 2)
div =           csv_dic('i/rebate_mdf/auxiliary/divs.csv')
gl_parent =     csv_dic('i/rebate_mdf/auxiliary/gl_parent.csv', 2)
job_numbers =   csv_dic('i/rebate_mdf/auxiliary/job_numbers.csv', 4)
vendors =       csv_dic('i/rebate_mdf/auxiliary/vendors.csv')

############### utils ###############
    
def scan_oracle(r, actual_plan, year, qtr):
    r.pop('', None)
    r['Actual / Plan'] = ap.get(actual_plan, '')
    r['Amount'] = r['Amount'].replace(',', '')
    r['Amount'] = float(r['Amount']) * -1
    r['Department Desc'] = departments.get(r['Department'], r['Department'])[0]
    r['Division'] = div.get(r['Division'], r['Division'])
    if r['GL Account'] in gl_parent:
        r['GL Parent'] = gl_parent[r['GL Account']][1]
    else:
        r['GL Parent'] = ''
    r['Quarter'] = qtr
    if r['Category'] in categories:
        r['SCC Category'] = categories[r['Category']][0]
        r['Super Category'] = categories[r['Category']][1]
        r['Super Category 2'] = categories[r['Category']][2]
    else:
        r['SCC Category'] = 'Corporate'
        r['Super Category'] = 'Corporate'
        r['Super Category 2'] = 'Corporate'
    r['Vendor'] = vendors.get(r['Vendor'], 'All Other')
    r['Year'] = year

    # Non-2015 Cisco Super Category
    if (('CISCO' in r['Vendor']) and (r['Category'] not in ['421', '422', '425']) and
        r['Year'] != 2015):
        r['SCC Category'] = 'Cisco'
        r['Super Category'] = 'Cisco'  
        r['Super Category 2'] = 'Cisco'
    
    # Job Number Assignment
    if (r['Project'][:3] in job_numbers or r['Project'] in job_numbers):
        r['Corporate or Custom'] = job_numbers[r['Project']][0]
        r['Marketing Category'] = job_numbers[r['Project']][1]
        r['Marketing Sub Category'] = job_numbers[r['Project']][2]
        r['Marketing Details'] = r['Project'] + ' - ' + job_numbers[r['Project']][3]
    else:
        r['Corporate or Custom'] = 'Default'
        r['Marketing Category'] = 'Default'
        r['Marketing Sub Category'] = 'Default'
       
    if r['Amount'] != 0:
        rows.append(r)

def scan_csv():
    for file in os.listdir(input_dir):
        file_path = input_dir + file

        # Oracle - Rebate / MDF
        if file.endswith(".csv"):
            actual_plan, year, qtr = file.split('_')
            qtr = qtr[:2]
            input_file = csv.DictReader(open(file_path))
            for r in input_file:
                scan_oracle(r, actual_plan, year, qtr)
            print file

def write_csv():
    headers = ['Actual / Plan', 'Amount', 'BD Dept', 'Corporate or Custom', 
               'Category', 'Description', 'Department', 'Department Desc', 'Division', 
               'GL Parent', 'Marketing Category', 'Marketing Sub Category', 
               'Marketing Details', 'Project', 'Quarter', 'SCC Category', 
               'Solution Group', 'Solution Type', 'Super Category', 
               'Super Category 2', 'Territory', 'Vendor', 'Year']

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
