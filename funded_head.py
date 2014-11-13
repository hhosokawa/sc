from aux_reader import *
import csv
import os
from pprint import pprint
import time

# io
input_dir = 'i/funded_head/'
output = 'o/funded_head.csv'
rows = []

# Pictionary
departments =   csv_dic('i/funded_head/auxiliary/departments.csv', 2)
div =           csv_dic('i/funded_head/auxiliary/divs.csv')
gl_parent =     csv_dic('i/funded_head/auxiliary/gl_parent.csv', 2)

############### utils ###############

def scan_oracle(r, year, period):
    r.pop('', None)
    r['Amount'] = r['Amount'].replace(',', '')
    r['Amount'] = float(r['Amount']) * -1
    r['Division'] = div.get(r['Division'], r['Division'])
    if r['Department'] in departments:
        r['Department A'] = departments[r['Department']][1]
        r['Department B'] = departments[r['Department']][0]
    else:
        r['Department A'] = 'Corporate'
        r['Department B'] = 'Corporate'
    r['Period'] = period
    r['Year'] = year
      
    if r['Amount'] != 0:
        rows.append(r)

def scan_csv():
    for file in os.listdir(input_dir):
        file_path = input_dir + file

        # Oracle - Rebate / MDF
        if file.endswith(".csv"):
            year, period = file.split('_')
            period = int(period[1:3])
            input_file = csv.DictReader(open(file_path))
            for r in input_file:
                scan_oracle(r, year, period)
            print file

def write_csv():
    headers = ['Amount', 'Description', 'Department', 'Department A', 
               'Department B', 'Division', 'Period', 'Year']

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
    print 'funded_head.py complete.', t1-t0
