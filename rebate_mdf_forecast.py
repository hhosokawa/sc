import csv
import os
import pandas as pd
from pprint import pprint
import time

# Input / Output
INPUT_DIR = 'i/rebate_mdf_forecast/'
ZERO_ROWS = 'i/rebate_mdf_forecast/auxiliary/zero_rows.csv'
OUTPUT = 'o/rebate_mdf_forecast/rebate_mdf_data.csv'

# Converts csv -> Dictionary
def csv_dic(filename, style = 1):     # Converts CSV to dict
    reader = csv.reader(open(filename, "rb"))
    if style == 1: 
        my_dict = dict((k, v) for k, v in reader)
    elif style == 2:
        my_dict = dict((k, (v1, v2)) for k, v1, v2 in reader)
    return my_dict

# Auxiliary dictionaries
categories = csv_dic('i/rebate_mdf_forecast/auxiliary/categories.csv', 2)
divisions = {'100':'Canada', '200':'United States'}
plans = {'A':'Actual', 'B':'Plan'}
gl_parent = csv_dic('i/rebate_mdf_forecast/auxiliary/gl_parent.csv')
territories = csv_dic('i/rebate_mdf_forecast/auxiliary/territories.csv')
vendors = csv_dic('i/rebate_mdf_forecast/auxiliary/vendors.csv')

############### Data Cleanse Functions ###############
# Oracle - Rebate and MDF Revenue / Expense
def clean_oracle(r, plan, book, period, year):
    
    # Clean Formatting
    r.pop('', None)
    if ',' in r['Amount']: 
        r['Amount'] = r['Amount'].replace(',', '')
    r['Book'] = book
    r['Year'] = year
    r['Actual-Plan'] = plans.get(plan, None)
    r['Division'] = divisions.get(r['Division'], r['Division'])
    r['GL Parent'] = gl_parent.get(r['GL Account'], None)
    r['Period'] = period
    r['Vendor Name'] = vendors.get(r['Vendor'], r['Vendor'])
    r['Region'] = territories.get(r['Territory'], r['Territory'])
    
    # Fix Corporate region -> Canada / US
    if r['Region'] == 'Corporate':
        if r['Division'] == 'Canada':
            r['Region'] = 'Canada Region'
        else:
            r['Region'] = 'US Corporate'
    r['SCC Category'], r['Super Category'] = categories.get(r['Category'], None)

    # Assign Quarter
    if 1 <= period <= 3:       r['Quarter'] = 1
    elif 4 <= period <= 6:     r['Quarter'] = 2
    elif 7 <= period <= 9:     r['Quarter'] = 3
    elif 10 <= period <= 12:   r['Quarter'] = 4        
        
    # If Amount != 0, include in rows
    if float(r['Amount']) != 0:
        r['Amount'] = float(r['Amount']) * -1
        return r
    
############## Data Output ###############
def scan_csv():
    headers = sorted(['Amount', 'Book', 'Year', 'Category', 'Department', 'Description', 
               'Division', 'Period', 'Quarter', 'GL Account', 'GL Parent', 
               'Vendor Name', 'Region', 'SCC Category', 'Super Category', 
               'Territory', 'Actual-Plan', 'Project'])

    with open(OUTPUT, 'wb') as o0:
        o0w = csv.DictWriter(o0, delimiter=',',
                             fieldnames=headers, extrasaction='ignore')
        o0w.writerow(dict((fn, fn) for fn in headers))
 
        # Scan *.csv input directory - Data Cleanse
        for file in os.listdir(INPUT_DIR):
            file_path = INPUT_DIR + file

            # Oracle - Rebate and MDF Revenue / Expense
            if file.endswith('.csv'):
                plan, book, period, year = file.split('_')
                period = int(period)
                year = int(year[:-4])    # remove .csv extension
                input_file = csv.DictReader(open(file_path))
                for r in input_file:
                    r = clean_oracle(r, plan, book, period, year)
                    if r: o0w.writerow(r)
    print 'write_csv() complete.'
    
# Create 0 dollar rows
def append_zero_rows():
    df = pd.read_csv(ZERO_ROWS, header=None)
    with open(OUTPUT , 'a') as f:
        df.to_csv(f, header=False, index=False)
    print 'append_zero_rows() complete.'

############### main ###############
if __name__ == '__main__':
    t0 = time.clock()
    scan_csv()
    append_zero_rows()
    t1 = time.clock()
    print 'rebate_mdf_forecast.py complete.', t1-t0    