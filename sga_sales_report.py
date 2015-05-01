from aux_reader import *
from datetime import datetime
from dateutil.parser import parse
from openpyxl import Workbook
from pprint import pprint
import os
import time

# io
input_dir = 'i/sga_sales_report/'       # Update
output = 'o/sga_sales_report.xlsx'  
headcount_dict = {}
rows = []
supplier_name_rows = []

# Pictionary
ap = {'A':'Actual', 'B':'Plan'}
departments = csv_dic('i/sga_sales_report/auxiliary/departments.csv', 2)
headcount = csv_dic('i/sga_sales_report/auxiliary/headcount.csv', 6)    # Update
hierarchies = csv_dic('i/sga_sales_report/auxiliary/hierarchies.csv', 2)
territories = csv_dic('i/sga_sales_report/auxiliary/territories.csv', 3)

############### Data Extract ###############
# Scan Input Directory
def scan_csv():
    for file in os.listdir(input_dir):
        file_path = input_dir + file

        # A) BI - Field Margin
        if (file == 'bi_fm_A.csv') or (file == 'bi_fm_B.csv'):
            _, _, actual_plan = file.split('_')
            actual_plan = actual_plan[:1]
            input_file = csv.DictReader(open(file_path))
            for r in input_file:
                clean_bi(r, actual_plan)
                
        # B) Discoverer - Employee Expense Audit Report
        elif file == 'Employee Expense Audit Report.csv':
            input_file = csv.DictReader(open(file_path))
            for r in input_file:
                clean_expense_report(r)
                
        # C) Discoverer - Supplier Name Sales Expense
        elif file == 'sales_report_supplier_name.csv':
            input_file = csv.DictReader(open(file_path))
            for r in input_file:
                clean_supplier_name_report(r)

        # D) Oracle - SG & A Expenses
        elif file.endswith(".csv"):
            year, period, book, currency, actual_plan = file.split('_')
            actual_plan = actual_plan[:1]
            input_file = csv.DictReader(open(file_path))
            for r in input_file:
                clean_oracle(r, year, period, book, currency, actual_plan)
    print 'scan_csv() complete.'

# A) BI - Field Margin
def clean_bi(r, actual_plan):
    # Assign correct headers
    r['Actual/Plan'] = ap[actual_plan]
    r['GL GrandParent'] = 'Field Margin'
    r['GL Parent'] = 'Field Margin'
    r['Period'] = r['Fiscal Period']
    r['Quarter'] = r['Fiscal Quarter']
    r['Year'] = r['Calendar Year']

    # Clean CA Corporate -> Canada
    if ((r.get('Region','') == 'CA Corporate') 
    or (r.get('OB/TSR Region', '') == 'CA Corporate')):
        r['Region'] = 'Canada'
        r['District'] = 'CA Corporate'
        
    # If 'OB / TSR' == blank or 'IS' -> 'TSD'
    if r.get('OB or TSR', '') in ['','IS']:
        r['OB or TSR'] = 'TSD'
        
    # Actuals
    if r['Actual/Plan'] == 'Actual':

        if 'Region' in r['Region']:
            r['Region'] = r['Region'].replace('Region', '')
            r['Region'] = r['Region'].strip()

        # Clean SC Canada / SC United States
        r['Amount'] = r['Virtually Adjusted GP']
        if r['Division'] == 'Canada':
            r['Book'] = 'SC Canada'
            r['Currency'] = 'CAD'
        else:
            r['Book'] = 'SC United States'
            r['Currency'] = 'USD'

        if r['Amount'] != 0:
            rows.append(r)

    # Plan
    elif r['Actual/Plan'] == 'Plan':

        # Clean SC Canada / SC United States
        r['Amount'] = r['Branch GP Plan']
        if r['Division'] == 'Canada':
            r['Book'] = 'SC Canada'
            r['Currency'] = 'CAD'
        else:
            r['Book'] = 'SC United States'
            r['Currency'] = 'USD'

        # Clean Region / District
        r = clean_plan_region_district(r)

        if r:
            if r['Amount'] != 0:
                rows.append(r)

# A) BI - Clean BI Plan Region / District
def clean_plan_region_district(r):
    district = r['District']
    try:
        # Remove Region text
        if 'Region' in r['Region']:
            r['Region'] = r['Region'].replace('Region', '')
            r['Region'] = r['Region'].strip()
            
        if district in ['CA Corporate', 'CA SMB District']:
            r['District'] = 'CA Corporate'
            r['Region'] = 'Canada'
        elif district in ['Atlanta Corporate Center', 'Chicago Corporate Center',
                          'Inactive (US Branches)', 'US Corporate']:
            r['Region'] = 'Corporate'
        return r
    except TypeError:
        pass

# B) Oracle - Spreadsheet Server
def clean_oracle(r, year, period, book, currency, actual_plan):

    # Clean Formatting, Assign GL, Region Hierarchies
    r.pop('', None)
    r['Actual/Plan'] = ap[actual_plan]
    if ',' in r['Amount']: 
        r['Amount'] = r['Amount'].replace(',', '')
    r['Amount'] = float(r['Amount'])
    r['Book'] = book
    r['Currency'] = currency
    if r['Department'] in departments:
        r['OB or TSR'] = departments[r['Department']][1]
    else:
        r['OB or TSR'] = 'Corporate'
    r['GL Parent'] = hierarchies.get(r['GL Account'], None)[1]
    r['Period'] = int(period)
    if 1 <= r['Period'] <= 3:       r['Quarter'] = 1
    elif 4 <= r['Period'] <= 6:     r['Quarter'] = 2
    elif 7 <= r['Period'] <= 9:     r['Quarter'] = 3
    elif 10 <= r['Period'] <= 12:   r['Quarter'] = 4
    divloc = r['Territory']
    r['Territory'], r['District'], r['Region'] = territories.get(divloc, None)
    r['Year'] = year
    
    # Determine if row is a discretionary expense
    if r['GL Parent'] in ['Meal Expense', 'Sales Expense', 'Travel & Entertainment']:
        r['Discretionary Expense'] = True
        r['GL GrandParent'] = 'Sales Expense'
    else:
        r['Discretionary Expense'] = False
    
    # If criteria met, append to output
    if ((r['Amount'] != 0) and 
        (r['Book'] != 'SC Consol - US') and
        (r['Discretionary Expense'] == True)):
        rows.append(r)

# C) Discoverer - Expense Report clean up
def clean_expense_report(r):
    if r['Employee Number'] in headcount:
        department = headcount[r['Employee Number']][1]
        
        # Collect only Outside Sales / Telesales Expense Reports
        if ('Outside Sales' in department) or ('Telesales' in department):
            unique_id = r['Employee Number'] + r['Invoice Num']
            if unique_id not in headcount_dict:
                divloc = headcount[r['Employee Number']][5]
                r['District'] = territories[divloc][1]
                r['Functional Group'] = headcount[r['Employee Number']][1]
                r['Region'] = territories[divloc][2]
                r.pop('Amount COUNT', None)
                headcount_dict[unique_id] = r
                
# D) Discoverer - Supplier Name Expense Report clean up
def clean_supplier_name_report(r):
    if r['GL Account'] in hierarchies:
        r['GL Description'] = hierarchies[r['GL Account']][0]
        r['GL Parent'] = hierarchies[r['GL Account']][1]
    
        # Determine if row is a discretionary expense
        if r['GL Parent'] in ['Meal Expense', 'Sales Expense', 'Travel & Entertainment']:
            create_date = parse(r['Creation Date'])
            past_date = datetime(2015,01,01)
            present_date = datetime.now()
            
            # Determine if expense is within current year
            if present_date >= create_date >= past_date:
                
                # Determine if OB / Telesales related expense
                if r['Department'] != '0000' and r['Department'] in departments:
                    r['District'] = territories.get(r['Territory'], '000')[1]
                    r['Period Name'] = parse(r['Creation Date']).month
                    r['Region'] = territories.get(r['Territory'], '000')[2]
if float(r['Line Amount SUM']) != 0:
                        supplier_name_rows.append(r)

# Create new hierarchy - Category A-D
def create_hierarchy():
    rows_copy = []
    for r in rows:
        if (r['OB or TSR'] in ['TSD', 'Telesales', 'TSR'] and 
            r['Region'] != 'US Govt'):
            r['Category A'] = 'Telesales / TSD / Corporate'
            r['Category B'] = r['OB or TSR']
            r['Category C'] = r['GL Parent']
            r['Category D'] = r['GL Parent']
        elif r['Region'] in ['Corporate', 'US Corporate']:
            r['Category A'] = 'Telesales / TSD / Corporate'
            r['Category B'] = 'Corporate'
            r['Category C'] = r['GL Parent']
            r['Category D'] = r['GL Parent']
        else:
            r['Category A'] = 'Regions'
            r['Category B'] = r['Region']
            r['Category C'] = r['District']
            r['Category D'] = r['GL Parent']
        rows_copy.append(r)
    print 'create_hierarchy() complete.'
    return rows_copy

############### Data Output ###############
def write_xlsx():
    pprint(rows[:5])
    print '\n'
    pprint(supplier_name_rows[:5])
    print '\n'
    for x in headcount_dict:
        print headcount_dict[x]
        raw_input()

############### main ###############
if __name__ == '__main__':
    t0 = time.clock()
    scan_csv()
    rows = create_hierarchy()
    write_xlsx()
    t1 = time.clock()
    print 'sga_sales_report.py complete.', t1-t0
