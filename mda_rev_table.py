import csv
import time
from aux_reader import *
from itertools import product

""" io """

output = 'o/mda_rev_table.csv'
gls = csv_dic('i/mda_rev_table/gl.csv')
categories = csv_dic('i/mda_rev_table/categories.csv')

rows = []
months =    range(1,13)
divs =      ['"*"', '"100"', '"200"']
years =     ['2014', '2013', '2012', '2011', '2010']
div_book =  {'"200"' : 'US', '"100"' : 'Canada', '"*"' : 'Consolidated'}

quarterperiod = {1: 'Q1', 2: 'Q1', 3: 'Q1',
                 4: 'Q2', 5: 'Q2', 6: 'Q2',
                 7: 'Q3', 8: 'Q3', 9: 'Q3',
                 10:'Q4', 11:'Q4', 12:'Q4'}

############## utils ###############

# Make Spreadsheet Server Formula
def makeformula(year, month, acct, cat, div):
    return ('=-' +
            'GXL("A","","TRANSLATED="&"E"&";"&"CURRENCY="&' +
            '"USD"' + '&";"&"BOOK="&' +     # Currency
            '"SC Consol - US"' + '&";",' +  # Book
            year + ',' +                    # Year
            '"PER"' + ',' +                 # PER or LTD
            str(month) + ',' +              # Period
            div + ',' +                     # Div
            '"*"' + ',' +                   # Dept
            '"*"' + ',' +                   # Territory
            acct + ',' +                    # GL Acct
            cat + ',"*")')                  # Category

# Concat MS Agency Fee Formula
def concat_ms_fee(year, month, cat, div, formula):
    ms_fees_formula = str(makeformula(year,
                          month, '"631000"', cat, div))
    formula = formula + ' + ' + ms_fees_formula[2:]
    return formula

def generate_rows():
    for acct, year, cat, month, div in product(gls, years, categories, months, divs):
        desc = gls[acct]
        div_desc = div_book[div]
        cat_desc = categories[cat]
        quarter = quarterperiod[month]
        formula = str(makeformula(year, month, acct, cat, div))

        # Net Sales Edit - (Revenue - MS Agency Fees)
        if desc == 'Net Sales':
            formula = concat_ms_fee(year, month,  cat, div, formula)

        r = (acct, desc, year, month, quarter,
             div_desc, cat_desc, formula)
        rows.append(r)
    print 'generate_rows() complete.'

def write_csv():
    headers = ['GL', 'GL Desc', 'Year', 'Month',
               'Quarter', 'Division', 'Super Category', 'FM']

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