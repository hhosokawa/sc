import csv
import time
from aux_reader import *
from itertools import product

""" io """

output = 'o/rebate_mdf.csv'
gls = csv_dic('i/rebate_mdf/gl.csv')
vendors = csv_dic('i/rebate_mdf/vendors.csv')
categories = csv_dic('i/rebate_mdf/categories.csv')

rows = []
qtrs =      range(1,5)
years =     range(2012, 2015)
divs =      {'"200"' : 'US',
             '"097"' : 'US',
             '"100"' : 'Canada',
             '"*"' : 'Consolidated'}

############## utils ###############

# Make Spreadsheet Server Formula
def makeformula(year, qtr, acct, cat, div, vendor):
    return ( '=-' +
            'GXL("A","","TRANSLATED="&"E"&";"&"CURRENCY="&' +
            '"USD"' + '&";"&"BOOK="&' +     # Currency
            '"SC Consol - US"' + '&";",' +  # Book
            str(year) + ',' +               # Year
            '"QTR"' + ',' +                 # PER / LTD / QTR
            str(qtr) + ',' +                # Period / Quarter
            div + ',' +                     # Div
            '"*"' + ',' +                   # Dept
            '"*"' + ',' +                   # Territory
            acct + ',' +                    # GL Acct
            cat + ',' +                     # Category
            '"*"' + ',' +                   # Job Number
            vendor + ')')                    # Vendor

def generate_rows():
    for acct, year, cat, qtr, div, vendor in product(gls, years, categories,
                                                     qtrs, divs, vendors):
        desc = gls[acct]
        div_desc = divs[div]
        cat_desc = categories[cat]
        vendor_desc = vendors[vendor]
        formula = str(makeformula(year, qtr, acct, cat, div, vendor))
        r = (acct, desc, year, qtr, div_desc, cat_desc, vendor_desc, formula)
        rows.append(r)
    print 'generate_rows() complete.'

def write_csv():
    headers = ['GL', 'GL Desc', 'Year', 'Quarter', 'Division',
               'Super Category', 'Vendor', 'Amount']

    with open(output, 'wb') as o1:
        o1w = csv.writer(o1)
        o1w.writerow(headers)
        for row in rows:
            o1w.writerow(row)
    print 'write_csv() complete.'

############## mda_rev_table_main() ###############

if __name__ == '__main__':
    t0 = time.clock()
    generate_rows()
    write_csv()
    t1 = time.clock()
    print 'mda_rev_table_main() completed. Duration:', t1-t0
