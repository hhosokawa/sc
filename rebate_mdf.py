from aux_reader import *
import csv
from itertools import product
import time

""" io """

output = 'o/rebate_mdf.csv'
#cats = csv_dic('i/rebate_mdf/categories.csv')
cats = {'"*"': 'Consolidated'}
gls = csv_dic('i/rebate_mdf/gl.csv')

rows = []
qtrs = range(1,5)
years = range(2014, 2015)
books = {'"SC Canada"' : ('Canada', '"M"'),
         '"SC Consol - US"' : ('Consolidated', '"E"'),
         '"SC United States"' : ('US', '"E"')}

############## utils ###############

def edit_formula(translate, book, year, qtr, gl, cat, desc, formula):

	# Operating Expense Edit (- Rebate Expense)
    if desc == 'Operating Expenses':
        rebate_expense_formula = str(makeformula(translate, book,
                                     year, qtr, '"^^000871_GL_Account"',cat))
        formula = formula + ' + ' + rebate_expense_formula[2:]

    # FM = Net Revenue - COGS - Rebate Revenue - MDF Revenue
    elif desc == 'FM':
        cogs_formula = str(makeformula(translate, book, year, qtr,
                                       '"^^000007_GL_Account"', cat))
        mdf_rev_formula  = str(makeformula(translate, book, year, qtr,
                                       '"^^000795_GL_Account"', cat))
        rebate_formula = str(makeformula(translate, book, year, qtr,
                                       '"^^000791_GL_Account"', cat))
        formula = (formula + ' - ' + cogs_formula[2:] + ' + ' +
                   rebate_formula[2:] + ' + ' +  mdf_rev_formula[2:])
    return formula

# Make Spreadsheet Server Formula
def makeformula(translate, book, year, qtr, gl, cat):
    return ('=-' +
            'GXL("A","","TRANSLATED="&' +
            translate +'&";"&"CURRENCY="&'+ # Translated (T,E,M)
            '"USD"' + '&";"&"BOOK="&' +     # Currency
            book + '&";",' +                # Book
            str(year) + ',' +               # Year
            '"QTR"' + ',' +                 # PER / LTD / QTR
            str(qtr) + ',' +                # Period / Quarter
            '"*"' + ',' +                   # Div
            '"*"' + ',' +                   # Dept
            '"*"' + ',' +                   # Territory
            gl + ',' +                      # GL Acct
            cat + ',' +                     # Category
            '"*"' + ',' +                   # Job Number
            '"*"' + ')')                    # Vendor

def generate_rows():
    for book, year, qtr, gl, cat in product(books, years, qtrs,
                                            gls, cats):
        book_desc, translate = books[book]
        cat_desc = cats[cat]
        desc = gls[gl]
        formula = str(makeformula(translate, book, year, qtr, gl, cat))
        formula = edit_formula(translate, book, year, qtr,
                               gl, cat, desc, formula)
        r = (gl, desc, year, qtr, book_desc, cat_desc, formula)
        rows.append(r)
    print 'generate_rows() complete.'

def write_csv():
    headers = ['GL', 'GL Desc', 'Year', 'Quarter', 'Division',
               'Super Category', 'Amount']

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
