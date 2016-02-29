import csv
import os
from pprint import pprint
import time

# Input / Output
INPUT_DIR = 'i/services_oracle_ss/'
OUTPUT = 'o/services_oracle_ss.csv'

# Converts csv -> Dictionary
def csv_dic(filename, style = 1):     # Converts CSV to dict
    reader = csv.reader(open(filename, "rb"))
    if style == 1: 
        my_dict = dict((k, v) for k, v in reader)
    elif style == 2:
        my_dict = dict((k, (v1, v2)) for k, v1, v2 in reader)
    return my_dict

# Dictionaries
categories = csv_dic('i/services_oracle_ss/dictionaries/categories.csv', 2)
divisions = {'100':'Canada', '200':'United States'}
vendors = csv_dic('i/services_oracle_ss/dictionaries/vendors.csv')

# Gross Revenue Arrays GLs
custom_gross_rev_array = ['622100', '685100', '612150', '612210', '665200',
                          '612140', '612130', '612120', '612110', '665100']
gross_rev_array = ['612110', '612120','612130', '612140', '612150',
                   '612210', '622100', '712100', '712110', '712120',
                   '712200', '712300', '722100']
rebate_array = ['795110']
mdf_array = ['795210', '795220', '795230', '795240']

############### Data Cleanse Functions ###############
# Clean Oracle
def clean_oracle(r, period, year):
    
    # Clean Formatting
    r.pop('', None)
    if ',' in r['Amount']: 
        r['Amount'] = r['Amount'].replace(',', '')
    r['Division'] = divisions.get(r['Division'], r['Division'])
    r['Fiscal Period'] = period
    r['Managed Vendor Name'] = vendors.get(r['Vendor'], r['Vendor'])
    r['Year'] = year
    if r['Category'] in categories:
        r['SCC Category'], r['Super Category'] = categories[r['Category']]
    else:
        r['SCC Category'], r['Super Category'] = r['Category'], r['Category']
    # Assign Quarter
    if 1 <= period <= 3:       r['Fiscal Quarter'] = 1
    elif 4 <= period <= 6:     r['Fiscal Quarter'] = 2
    elif 7 <= period <= 9:     r['Fiscal Quarter'] = 3
    elif 10 <= period <= 12:   r['Fiscal Quarter'] = 4        
        
    # If Amount != 0, include in rows
    if float(r['Amount']) != 0:
        r['Amount'] = float(r['Amount']) * -1
        
        # Isolate FM, GP, Gross Revenue, MDF Revenue, Net Revenue, Rebate
        if r['Project'] != '200351':
            r['USD GP'] = r['Amount']                           # GP
            if r['GL Account'] not in rebate_array + mdf_array: # FM
                r['USD FM'] = r['Amount']       
            if ((r['GL Account'] not in gross_rev_array)        # Gross Revenue
                and (r['GL Account'][:1] == '6')
                and (r['GL Account'] not in custom_gross_rev_array)):  
                r['USD Gross Revenue'] = r['Amount']
            if r['GL Account'] in rebate_array:                 # Rebate
                r['USD Rebate'] = r['Amount']
            if r['GL Account'] in mdf_array:                    # MDF Revenue
                r['USD MDF Revenue'] = r['Amount']
            if r['GL Account'][:1] == '6':                      # Net Revenue
                r['USD Net Revenue'] = r['Amount']
        return r

############### Data Output ###############
def scan_csv():
    headers = sorted(['Amount', 'Year', 'Category', 'Department', 'Description', 'Division', 
               'Fiscal Period', 'Fiscal Quarter', 'GL Account', 'Managed Vendor Name', 
               'SCC Category', 'Super Category', 'USD FM', 'USD Gross Revenue', 'USD GP', 
               'USD Rebate', 'USD MDF Revenue', 'USD Net Revenue', 'Vendor',
               'Project'])

    with open(OUTPUT, 'wb') as o0:
        o0w = csv.DictWriter(o0, delimiter=',',
                             fieldnames=headers, extrasaction='ignore')
        o0w.writerow(dict((fn, fn) for fn in headers))
 
        # Scan *.csv input directory - Data Cleanse
        for file in os.listdir(INPUT_DIR):
            file_path = INPUT_DIR + file
                    
            if (file.endswith('.csv')):
                book, period, year = file.split('_')
                year = int(year[:-4])    # remove .csv extension
                input_file = csv.DictReader(open(file_path))
                for r in input_file:
                    r = clean_oracle(r, period, year)
                    if r: o0w.writerow(r)
    print 'write_csv() complete.'

############### main ###############
if __name__ == '__main__':
    t0 = time.clock()
    scan_csv()
    t1 = time.clock()
    print 'services_oracle_ss.py complete.', t1-t0    