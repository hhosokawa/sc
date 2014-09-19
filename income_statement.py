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
categories =    csv_dic('i/income_statement/auxiliary/categories.csv', 2)
departments =   csv_dic('i/income_statement/auxiliary/departments.csv', 2)
divs =          csv_dic('i/income_statement/auxiliary/divs.csv')
gl_parent =     csv_dic('i/income_statement/auxiliary/gl_parent.csv', 2)
territories =   csv_dic('i/income_statement/auxiliary/territories.csv')
vendors =       csv_dic('i/income_statement/auxiliary/vendors.csv')

############### utils ###############

# Assign GL B = Net Sales, Cost of Sales, Operating Income
def assign_GLB(r):

    # Net Sales -> Super Categor
    if r['GL A'] in ['Net sales', 'Cost of Sales']:
        return categories.get(r['Category'], 'Corporate')[1]

    # Operating Income -> Selling, Admin, Marketing, Amortization
    elif r['GL A'] == 'Operating expenses':
        glb = departments.get(r['Department'], r['Department'])[1]
        marketing_expenses = ['871100', '871110', '871120', '871130', '871140',
                              '871150', '871160', '871999']
        amortization_expenses = ['881000', '883000','891000','892000', '893000']
        if r['GL Account'] in marketing_expenses:
            glb = 'Marketing Expenses'
        elif r['GL Account'] in amortization_expenses:
            glb = 'Depreciation & Amortization'
        return glb
    else:
        return r['GL A']


def assign_region(r):
    region = territories.get(r['Territory'], r['Territory'])
    if region == 'Corporate':
        region = divs.get(r['Division'], r['Division'])
    return region

def scan_oracle(r, year, qtr):
    r.pop('', None)
    r['Amount'] = r['Amount'].replace(',', '')
    r['Amount'] = float(r['Amount']) * -1
    r['Category Desc'] = categories.get(r['Category'], 'Corporate')[0]
    #r['Division'] = divs.get(r['Division'],r['Division'])
    r['GL A'] = gl_parent.get(r['GL Account'], r['Description'])[1]
    #r['GL B'] = assign_GLB(r)
    r['Quarter'] = qtr
    #r['Super Category'] = categories.get(r['Category'], 'Corporate')[1]
    r['Region'] = assign_region(r)
    #r['Vendor'] = vendors.get(r['Vendor'], r['Vendor'])
    r['Year'] = year

    if r['Amount'] != 0:
        rows.append(r)

        # EBITDA Entry
        #if r['GL B'] not in ['Depreciation & Amortization', 
                             #'Other expenses (income)']:
            #r = r.copy()
            #r['GL A'] = 'EBITDA'
            #rows.append(r)
            

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
    headers = ['Amount', 'Category', 'Category Desc', 'GL Account', 'GL A', 
               'Department', 'Description', 'Division', 'Project', 'Region', 
               'Quarter', 'Territory', 'Year']

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
