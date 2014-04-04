import csv
import time
from aux_reader import *
from itertools import product

""" io """

output = 'o/mda_rev_table.csv'
gls = csv_dic('i/mda_rev_table/gl.csv')
categories = csv_dic('i/mda_rev_table/categories.csv')

rows = []
months = range(1,13)
years = ['2014', '2013', '2012', '2011']
quarterperiod = {1: 'Q1', 2: 'Q1', 3: 'Q1',
                 4: 'Q2', 5: 'Q2', 6: 'Q2',
                 7: 'Q3', 8: 'Q3', 9: 'Q3',
                 10:'Q4', 11:'Q4', 12:'Q4'}

############## utils ###############

# Make Spreadsheet Server Formula
def makeformula(year, month, acct, cat):
    return ('=-' +
            'GXL("A","","TRANSLATED="&"E"&";"&"CURRENCY="&' +
            '"USD"' + '&";"&"BOOK="&' +     # Currency
            '"SC Consol - US"' + '&";",' +  # Book
            year + ',' +                    # Year
            '"PER"' + ',' +                 # PER or LTD
            str(month) + ',' +              # Period
            '"*"' + ',' +                   # Div
            '"*"' + ',' +                   # Dept
            '"*"' + ',' +                   # Territory
            acct + ',' +                    # GL Acct
            cat + ',"*")')                  # Category

def generate_rows():
    for acct, year, cat, month in product(gls, years, categories, months):
        desc = gls[acct]
        cat_desc = categories[cat]
        quarter = quarterperiod[month]
        formula = str(makeformula(year, month, acct, cat))

        # Isolate Net Sales / MS Agency Fees
        if desc in ['Net Sales', 'MS Direct EAs']:
            rev_formula = formula
        # COGS
        else:
            rev_formula = 0

        # Edit Microsoft Desc and Revenue Fields
        if cat_desc == 'Microsoft':

            # Rename MS Net Sales / COGS -> MS License and SA
            if desc == 'Net Sales':
                desc = 'Net Sales - MS License and SA'
                ms_fees_formula = str(makeformula(year,
                                      month, '"631000"', cat))
                formula = formula + ' + ' + ms_fees_formula[2:]
                rev_formula = formula

        # Isolate CC / ESSN / Services
        else:

            # Ignore 'MS Direct EAs' for CC / ESSN / Services
            if desc == 'MS Direct EAs' and cat_desc != 'Services':
                continue

        r = (acct, desc, year, month, quarter,
             cat_desc, formula, rev_formula)
        rows.append(r)
    print 'generate_rows() complete.'

def write_csv():
    headers = ['GL', 'GL Desc', 'Year', 'Month',
               'Quarter', 'Super Category', 'FM', 'Rev']

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
