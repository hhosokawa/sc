from aux_reader import *
from datetime import datetime
from dateutil.parser import parse
from openpyxl import Workbook
from pprint import pprint
import os
import time

# io
input_dir = 'i/sales_expense_report_detail/'       # Update
output = 'P:/_HHOS/Sales Expense Reporting/data/oracle_sales_expense_detail.csv'
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

    if ',' in r['Invoice Distribution Functional Amount']:
        r['Amount'] = float(r['Invoice Distribution Functional Amount'].replace(',',''))
    else:
        r['Amount'] = float(r['Invoice Distribution Functional Amount'])
    if r['Department'] in departments:
        r['Department Desc'] = departments[r['Department']][1]
    else:
        r['Department Desc'] = ''        
    pprint(r)
    raw_input()
    
# Clean Oracle - Sales, Travel, Meal expenses
def clean_oracle(r, year, period, book, currency, actual_plan):

    if r['Category_'] != 'Purchase Invoices':
        r.pop('', None)
        if ',' in r['Base Credit']: 
            r['Base Credit'] = r['Base Credit'].replace(',', '')
        if ',' in r['Base Debit']: 
            r['Base Debit'] = r['Base Debit'].replace(',', '')
        r['Amount'] = float(r['Base Debit']) - float(r['Base Credit'])
        _, r['District'], r['Region'] = territories[r['Territory']]
        if r['Department'] in departments:
            r['Department Desc'] = departments[r['Department']][1]
        else:
            r['Department Desc'] = ''
        r['GL Parent'] = hierarchies[r['GL Account']][1]
        if r['Region'] == 'Corporate':
            if r['Division'] == '100':
                r['District'] = 'CA Corporate'
                r['Region'] = 'Canada'
            else:
                r['District'] = 'US Corporate'
                r['Region'] = 'US Corporate'
        if '.' in r['Line Description']:
            r['Line Description'] = r['Line Description'].replace('.', '')
        r['Supplier Name'] = employee_names.get(r['Line Description'],'')
        if r['Amount'] != 0:
            rows.append(r)

############### Data Output ###############
def write_csv():
    # Oracle - Sales Expense Report
    headers = ['Actual/Plan', 'Amount', 'Book', 'Currency', 'Category A',
               'Category B', 'Category C', 'Category D', 'Description',
               'Discretionary Expense', 'District', 'GL Account', 
               'GL GrandParent', 'GL Parent', 'OB or TSR', 'Period', 'Quarter', 
               'Region', 'Solution Group', 'Territory', 'Year']

    with open(output, 'wb') as o0:
        o0w = csv.DictWriter(o0, delimiter=',',
                             fieldnames=headers, extrasaction='ignore')
        o0w.writerow(dict((fn, fn) for fn in headers))
        for r in rows:
            o0w.writerow(r)
    
    # Discoverer - Supplier Name Report
    headers = ['Cat', 'Creation Date', 'Department', 'Division', 'GL Account',
               'GL Description', 'GL Parent', 'Invoice Date', 'Invoice Number',
               'Line Amount SUM', 'Period Name', 'Project', 'Region', 'Supplier Name',
               'Territory']

    with open(output1, 'wb') as o1:
        o1w = csv.DictWriter(o1, delimiter=',',
                             fieldnames=headers, extrasaction='ignore')
        o1w.writerow(dict((fn, fn) for fn in headers))
        for r in supplier_name_rows:
            o1w.writerow(r)
            
    # Discoverer - Employee Audit Expense Report + Headcount info
    headers = ['District', 'Employee Name', 'Employee Number', 
               'Functional Group',  'Invoice Num', 'Overrider Approver Name',
               'Receipts Received Date', 'Region', 'Report Submitted Date']    

    with open(output2, 'wb') as o2:
        o2w = csv.DictWriter(o2, delimiter=',',
                             fieldnames=headers, extrasaction='ignore')
        o2w.writerow(dict((fn, fn) for fn in headers))
        for id in headcount_dict:
            o2w.writerow(headcount_dict[id])
    print 'write_csv() complete.'
             
############### main ###############
if __name__ == '__main__':
    t0 = time.clock()
    scan_csv()
    rows = create_hierarchy()
    write_csv()
    t1 = time.clock()
    print 'sales_expense_report.py complete.', t1-t0
