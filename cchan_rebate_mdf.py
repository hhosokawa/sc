import csv
import time
from aux_reader import *
from itertools import product

""" io """

output = 'o/cchan_rebate_mdf.csv'
gls = csv_dic('i/cchan_rebate_mdf/gl.csv')
cad_us_fx = csv_dic('i/cchan_rebate_mdf/cad_us_fx.csv')

rows = []
pers =      range(1,13)
years =     range(2013, 2015)
divs =      {'"200"' : 'US',
             '"097"' : 'US',
             '"100"' : 'Canada',
             '"*"' :   'Consolidated'}

############## utils ###############

# Make Spreadsheet Server Formula
def makeformula(cur, book, year, per, div, acct):
    return ( '=-' +
            'GXL("A","","TRANSLATED="&"E"&";"&"CURRENCY="&' +
            cur + '&";"&"BOOK="&' +         # Currency
            book + '&";",' +                # Book
            str(year) + ',' +               # Year
            '"PER"' + ',' +                 # PER / LTD / QTR
            str(per) + ',' +                # Period / Quarter
            div + ',' +                     # Div
            '"*"' + ',' +                   # Dept
            '"*"' + ',' +                   # Territory
            acct + ',' +                    # GL Acct
            '"*"' + ',' +                   # Category
            '"*"' + ',' +                   # Job Number
            '"*"' + ')')                    # Vendor

def add_fx(formula, year, per):
    year_per_id = str(year) + str(per)
    formula += ' / ' + cad_us_fx.get(year_per_id, '1')
    return formula

def generate_rows():
    for acct, year, per, div in product(gls, years, pers, divs):
        desc = gls[acct]
        div_desc = divs[div]

        # Assign Currency and Book
        if div == '"100"':
            cur, book = '"USD"', '"SC Consol - US"'
        else:
            cur, book = '"USD"', '"SC Consol - US"'

        formula = str(makeformula(cur, book, year, per, div, acct))

        # CAD -> US FX
        if div == '"100"': adj_formula = add_fx(formula, year, per)
        else: adj_formula = formula
        r = (acct, desc, year, per, div_desc, formula, adj_formula)
        rows.append(r)
    print 'generate_rows() complete.'

def write_csv():
    headers = ['GL', 'GL Desc', 'Year', 'Period',
               'Division', 'Amount', 'Native Currency Amount']

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
