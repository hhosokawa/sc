import csv
import time
from aux_reader import *
from itertools import product

""" io """

output = 'o/mda_rev_table.csv'
gls = csv_dic('i/mda_rev_table/gl.csv')
categories = csv_dic('i/mda_rev_table/categories.csv')

rows = []
qtrs =      range(1,5)
years =     range(2009, 2015)
divs =      {'"200"' : 'US',
             '"097"' : 'US',
             '"100"' : 'Canada',
             '"*"' : 'Consolidated'}
job = '"[*,/200279,/200349,/200350,/200351]"'

############## utils ###############

# Make Spreadsheet Server Formula
def makeformula(year, qtr, acct, cat, div):
    return ('=-' +
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
            job + ')')                      # Job Number

def generate_rows():
    for acct, year, cat, qtr, div in product(gls, years, categories,
                                             qtrs, divs):
        desc = gls[acct]
        div_desc = divs[div]
        cat_desc = categories[cat]
        formula = str(makeformula(year, qtr, acct, cat, div))

        # Net Sales Edit - (Revenue - MS Agency Fees)
        if desc == 'Net Sales':
            ms_fees_formula = str(makeformula(year, qtr, '"631000"', cat, div))
            formula = formula + ' + ' + ms_fees_formula[2:]

        r = (acct, desc, year, qtr, div_desc, cat_desc, formula)
        rows.append(r)
    print 'generate_rows() complete.'

def write_csv():
    headers = ['GL', 'GL Desc', 'Year', 'Quarter',
               'Division', 'Super Category', 'FM']

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
