import csv
import time
from aux_reader import *
from itertools import product

""" io """

output = 'o/oracle_mc.csv'
gls = csv_dic('i/oracle/gl_parent.csv')
territories = csv_dic('i/oracle/territories.csv')
categories = csv_dic('i/oracle/categories.csv')

rows = []
years = ['2013', '2012', '2011']

############## utils ###############

# Make Spreadsheet Server Formula
def makeformula(year, month, ter, acct, cat):
    return ('=-' +
            'GXL("A","","TRANSLATED="&"E"&";"&"CURRENCY="&' +
            '"USD"' + '&";"&"BOOK="&' +     # Currency
            '"SC Consol - US"' + '&";",' +  # Book
            year + ',' +                    # Year
            '"LTD"' + ',' +                 # PER or LTD
            str(month) + ',' +              # Period
            '"*"' + ',' +                   # Div
            '"*"' + ',' +                   # Dept
            ter + ',' +                     # Territory
            acct + ',' +                    # GL Acct
            cat + ',"*")')                  # Category

def generate_rows():
    for acct, year, ter, cat in product(gls, years, territories, categories):
        desc = gls[acct]
        cat_desc = categories[cat]
        ter_desc = territories[ter]
        formula = makeformula(year, 12, ter, acct, cat)

        r = (acct, desc, year, ter_desc, cat_desc, str(formula))
        rows.append(r)
    print 'generate_rows() complete.'

def write_csv():
    headers = ['GL', 'GL Desc', 'Year', 'Region', 'Super Category', 'Amount']

    with open(output, 'wb') as o1:
        o1w = csv.writer(o1)
        o1w.writerow(headers)

        for row in rows:
            o1w.writerow(row)
    print 'write_csv() complete.'

############## oracle_main() ###############

if __name__ == '__main__':
    t0 = time.clock()
    generate_rows()
    write_csv()
    t1 = time.clock()
    print 'oracle_main() completed. Duration:', t1-t0
