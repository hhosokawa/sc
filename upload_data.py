import csv
import os
import time

# io
INPUT_DIR = 'i/upload_data/'
OUTPUT = 'o/upload_data.csv'

# Converts csv -> Dictionary
def csv_dic(filename, style = 1):     # Converts CSV to dict
    reader = csv.reader(open(filename, "rb"))
    if style == 1: 
        my_dict = dict((k, v) for k, v in reader)
    elif style == 2:
        my_dict = dict((k, (v1, v2)) for k, v1, v2 in reader)
    return my_dict

# Pictionary
rows = []
bi_vendors = csv_dic('i/upload_data/dictionaries/bi_vendors.csv')
categories = csv_dic('i/upload_data/dictionaries/categories.csv', 2)
divisions = {'100':'Canada', '200':'United States'}
gl_parent = csv_dic('i/upload_data/dictionaries/gl_parent.csv', 2)
gl_parent_reporting = csv_dic('i/upload_data/dictionaries/gl_parent_reporting.csv', 2)
territories = csv_dic('i/upload_data/dictionaries/territories.csv')
vendors = csv_dic('i/upload_data/dictionaries/vendors.csv')

############### Data Cleanse ###############
# Scan Input Directory
def scan_csv():
    for f in os.listdir(INPUT_DIR):
        file_path = INPUT_DIR + f
        
        # BI - Field Margin
        if f == 'bi.csv':
            input_file = csv.DictReader(open(file_path))
            for r in input_file:
                clean_bi(r)

        # Oracle - Rebate and MDF Revenue / Expense
        elif f.endswith(".csv"):
            year, period, book = f.split('_')
            book = book[:-4]    # remove .csv extension
            period = int(period)
            input_file = csv.DictReader(open(file_path))
            for r in input_file:
                clean_oracle(r, year, period, book)
    print 'scan_csv() complete.'
    
# Oracle - Rebate and MDF Revenue / Expense
def clean_bi(r):
    # Assign COGS
    r['COGS'] = float(r['Virtually Adjusted Revenue']) - float(r['Virtually Adjusted GP'])
    r['Amount'] = r['COGS']
    r['GL Parent (Reporting)'] = 'COGS'
    r['GL Parent (Sharon)'] = 'COGS'
    
    # Clean Super Category
    if (r['Solution Group'] == 'SERVICES') or (r['Super Category'] == 'Professional Services (INACTIVE)'):
        r['Super Category'] = 'Services'
    elif r['SCC Category'] in ['Enterprise Software', 'Security']:
        r['Super Category'] = 'Enterprise SW & Security'
    elif r['Solution Type'] == 'SaaS - Cloud':
        r['Super Category'] = 'SaaS'

    # Clean Managed Vendor Name
    if r['Managed Vendor Name'] in bi_vendors:
        r['Managed Vendor Name'] = bi_vendors.get(r['Managed Vendor Name'])

    # Assign Book Value
    if r['Division'] == 'Canada':
        r['Book'] = 'CAN'
    else:
        r['Book'] = 'US'
    r1 = r.copy()
    rows.append(r1)
    
    # Create Consolidated Division
    r['Book'] = 'CON'
    r2 = r.copy()
    rows.append(r2)

def clean_oracle(r, year, period, book):
    # Clean Formatting
    r.pop('', None)
    if ',' in r['Amount']: 
        r['Amount'] = r['Amount'].replace(',', '')
    r['Book'] = book
    r['Calendar Year'] = year
    r['Division'] = divisions.get(r['Division'], r['Division'])
    r['GL Parent (Sharon)'] = gl_parent.get(r['GL Account'], None)[1]
    r['GL Parent (Reporting)'] = gl_parent_reporting.get(r['GL Account'], None)[1]
    r['Fiscal Period'] = period
    if 1 <= period <= 3:       r['Fiscal Quarter'] = 1
    elif 4 <= period <= 6:     r['Fiscal Quarter'] = 2
    elif 7 <= period <= 9:     r['Fiscal Quarter'] = 3
    elif 10 <= period <= 12:   r['Fiscal Quarter'] = 4
    r['Managed Vendor Name'] = vendors.get(r['Vendor'], r['Vendor'])
    r['Region'] = territories.get(r['Territory'], r['Territory'])
    if r['Region'] == 'Corporate':
        if r['Division'] == 'Canada':
            r['Region'] = 'Canada Region'
        else:
            r['Region'] = 'US Corporate'
    if r['Category'] in categories:
        r['SCC Category'], r['Super Category'] = categories[r['Category']]
    else:
        r['SCC Category'], r['Super Category'] = r['Category'], r['Category']
        
    # Correct Canada - US East -> Canada - Canada Region
    if (r['Division'] == 'Canada') and (r['Region'] == 'US East Region'):
        r['Region'] = 'Canada Region'
        
    # Fix "All Other" Managed Vendor Name for Microsoft and Cisco
    if (r['Super Category'] == 'Microsoft') and (r['Managed Vendor Name'] == 'All Other'):
        r['Managed Vendor Name'] = 'MICROSOFT'
    elif (r['Super Category'] == 'Cisco') and (r['Managed Vendor Name'] == 'All Other'):
        r['Managed Vendor Name'] = 'CISCO SYSTEMS'
        
    # Fix "CISCO SYSTEMS" in non "Services"
    if ((r['Managed Vendor Name'] == 'CISCO SYSTEMS') 
        and (r['Super Category'] in ['Data Center', 'Enterprise SW & Security', 'Other'])):
        r['Super Category'] = 'Cisco'
        
    # Fix "Microsoft" Super Category with "0704" Dept -> "Other" Super Category
    if r['Super Category'] == 'Microsoft' and r['Department'] == '0704':
        r['Super Category'] = 'Other'

    # Create Number - Name Fields
    r['Category Number Name'] = str(r['Category']) + '-' + r['SCC Category']
    r['GL Number Name'] = str(r['GL Account']) + '-' + r['Description']
    r['Vendor Number Name'] = str(r['Vendor']) + '-' + r['Managed Vendor Name']
        
    # If Amount != 0, include in rows
    if float(r['Amount']) != 0:
        r['Amount'] = float(r['Amount']) * -1
        rows.append(r)

############### Data Output ###############
def write_csv():
    headers = sorted(['Calendar Year', 'Fiscal Quarter', 'Fiscal Period', 'Division',
                      'Region', 'Super Category', 'SCC Category', 'Managed Vendor Name',  
                      'Category', 'Department', 'Description', 'GL Account',
                      'GL Parent (Sharon)', 'Territory', 'Vendor', 'Amount', 'Book',
                      'GL Parent (Reporting)', 'Vendor Number Name', 'Category Number Name',
                      'GL Number Name', 'District', 'Project'])

    with open(OUTPUT, 'wb') as o0:
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
    print 'upload_data.py complete.', t1-t0