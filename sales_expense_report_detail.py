from aux_reader import *
from datetime import datetime
from dateutil.parser import parse
from openpyxl import Workbook
from pprint import pprint
import os
import time

# io
input_dir = 'i/sales_expense_report_detail/'       # Update
output = 'o/oracle_sales_expense_detail.csv'
rows = []

# Pictionary
ap = {'A':'Actual', 'B':'Plan'}
departments = csv_dic('i/sales_expense_report_detail/auxiliary/departments.csv', 2)
employee_names = csv_dic('i/sales_expense_report_detail/auxiliary/Employee Expense Audit Report.csv')
hierarchies = csv_dic('i/sales_expense_report_detail/auxiliary/hierarchies.csv', 2)
territories = csv_dic('i/sales_expense_report_detail/auxiliary/territories.csv', 3)

############### Data Cleanse ###############
# Scan Input Directory
def scan_csv():
    for file in os.listdir(input_dir):
        file_path = input_dir + file

        # Clean Payables Invoice Distribution
        if 'Payables' in file:
            input_file = csv.DictReader(open(file_path))
            for r in input_file:
                clean_payables(r)
                
        # Clean Oracle - Sales, Meal, Travel Expenses
        elif file.endswith(".csv"):
            year, period, book, currency, actual_plan = file.split('_')
            actual_plan = actual_plan[:1]
            input_file = csv.DictReader(open(file_path))
            for r in input_file:
                clean_oracle(r, year, period, book, currency, actual_plan)
    print 'scan_csv() complete.'

# Clean Payables
def clean_payables(r):

    # Amount
    if ',' in r['Invoice Distribution Functional Amount']:
        r['Amount'] = float(r['Invoice Distribution Functional Amount'].replace(',',''))
    else:
        r['Amount'] = float(r['Invoice Distribution Functional Amount'])
    # Department Description
    if r['Department'] in departments:
        r['Department Desc'] = departments[r['Department']][1]
    else:
        r['Department Desc'] = ''
    # District / Region via Territory
    _, r['District'], r['Region'] = territories[r['Territory']]
    if r['Region'] == 'Corporate':
        if r['Division'] == '100':
            r['District'] = 'CA Corporate'
            r['Region'] = 'Canada'
        else:
            r['District'] = 'US Corporate'
            r['Region'] = 'US Corporate'    
    # GL Description / Parent
    r['GL Description'] = hierarchies[r['GL Account']][0]
    r['GL Parent'] = hierarchies[r['GL Account']][1]
    # Period / Year
    dt = parse(r['Period Name'])
    r['Period'] = dt.month
    r['Year'] = dt.year
    if r['Amount'] != 0:
        rows.append(r)
    
# Clean Oracle - Sales, Travel, Meal expenses
def clean_oracle(r, year, period, book, currency, actual_plan):

    if r['Category_'] != 'Purchase Invoices':
        r.pop('', None)
        
        # Amount
        if ',' in r['Base Credit']: 
            r['Base Credit'] = r['Base Credit'].replace(',', '')
        if ',' in r['Base Debit']: 
            r['Base Debit'] = r['Base Debit'].replace(',', '')
        r['Amount'] = float(r['Base Debit']) - float(r['Base Credit'])
        # Department Description
        if r['Department'] in departments:
            r['Department Desc'] = departments[r['Department']][1]
        else:
            r['Department Desc'] = ''
        # District / Region via Territory
        _, r['District'], r['Region'] = territories[r['Territory']]
        if r['Region'] == 'Corporate':
            if r['Division'] == '100':
                r['District'] = 'CA Corporate'
                r['Region'] = 'Canada'
            else:
                r['District'] = 'US Corporate'
                r['Region'] = 'US Corporate'
        # GL Description / Parent
        r['GL Description'] = hierarchies[r['GL Account']][0]
        r['GL Parent'] = hierarchies[r['GL Account']][1]
        # Supplier name via Line Description
        if '.' in r['Line Description']:
            r['Line Description'] = r['Line Description'].replace('.', '')
        r['Supplier Name'] = employee_names.get(r['Line Description'], r['Line Description'])
        if r['Amount'] != 0:
            rows.append(r)

############### Data Output ###############
def write_csv():
    # Oracle - Sales Expense Report
    headers = ['Amount', 'Category', 'Category_', 'Department', 'GL Description',
               'Department Desc', 'Distribution Line Type Code', 'District', 
               'Division', 'GL Account', 'GL Parent', 'Invoice Description', 
               'Line Description', 'Period', 'Region', 'Source_', 
               'Supplier Name', 'Territory', 'Year']

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
    print 'sales_expense_report_detail.py complete.', t1-t0
