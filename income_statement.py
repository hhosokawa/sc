from aux_reader import *
import csv
import os
from pprint import pprint
import time

# io
input_dir = 'i/income_statement/'
output = 'o/income_statement.csv'
rows = []

# Pictionary
categories =    csv_dic('i/income_statement/auxiliary/categories.csv')
divs =           {'100':'Canada', '200':'US'}
gl_parent =     csv_dic('i/income_statement/auxiliary/gl_parent.csv', 2)
region =        csv_dic('auxiliary/div-region.csv')
vendors =       csv_dic('i/income_statement/auxiliary/vendors.csv')

############### utils ###############

def scan_oracle(r, year, qtr):
    r.pop('', None)
    r['Amount'] = r['Amount'].replace(',', '')
    r['Amount'] = float(r['Amount']) * -1
    try: div = r['IBM Chassis Referral - CBV Collections']
    except: div = r['Division']
    r['Division'] = divs.get(div,div)
    r['GL Parent'] = gl_parent.get(r['GL Account'], r['Description'])[1]
    r['Quarter'] = qtr
    r['Super Category'] = categories.get(r['Category'], 'Corporate')
    r['Vendor'] = vendors.get(r['Vendor'], r['Vendor'])
    r['Year'] = year

    if r['Amount'] != 0:
        rows.append(r)

def scan_csv():
    for file in os.listdir(input_dir):
        file_path = input_dir + file

        # Oracle - Rebate / MDF
        if file.endswith(".csv"):
            year, qtr = file.split('_')
            qtr = qtr[:2]
            input_file = csv.DictReader(open(file_path))
            for r in input_file:
                scan_oracle(r, year, qtr)
            print file

def write_csv():
    headers = ['Amount', 'Category', 'GL Account', 'GL Parent', 'Description', 
               'Division', 'Quarter', 'Super Category', 'Territory', 'Vendor', 
               'Year']

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
