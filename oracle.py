import csv
import time
from aux_reader import *
from itertools import product

""" io """

output = 'o\\oracle.csv'
input1 = 'i\\oracle\\gl.csv'

years = ['2013']
divs = ['"*"', '"100"', '"200"', '"300"',
        '"310"', '"320"', '"099"']

gl = {}
cat = {}
rows = []
quarterperiod = {1: 'Q1', 2: 'Q1', 3: 'Q1',
                 4: 'Q2', 5: 'Q2', 6: 'Q2',
                 7: 'Q3', 8: 'Q3', 9: 'Q3',
                 10:'Q4', 11:'Q4', 12:'Q4'}

div_book = {
    '"*"'   : ('"USD"', '"SC Consol - US"', '* - Consolidated'),
    '"098"' : ('"USD"', '"SC Consol - US"', '098 - Elimination Canada'),
    '"099"' : ('"USD"', '"SC Consol - US"', '099 - Elimination US'),
    '"100"' : ('"USD"', '"SC Consol - US"', '100 - Canada'),
    '"101"' : ('"CAD"', '"SC Canada - Goliath"', '101 - Canada Goliath'),
    '"200"' : ('"USD"', '"SC Consol - US"', '200 - US'),
    '"201"' : ('"USD"', '"SC Consol - US - Goliath"', '201 - US Goliath'),
    '"300"' : ('"USD"', '"SC Holdco US"', '300 - Holdco US'),
    '"310"' : ('"USD"', '"SC Holdings II"', '310 - Holding II'),
    '"315"' : ('"USD"', '"Goliath Acquisition"', '315 - Goliath'),
    '"320"' : ('"USD"', '"Softchoice Holdings"', '320 - Holding Cda'),
    '"325"' : ('"USD"', '"ULC Holdco"', '325 - ULC Holdco')}

############### aux ###############

# Make Spreadsheet Server Formula
def makeformula(minus, currency, book, year, per_ltd, month, div, acct):
    return ('='+ minus +'GXL("A","","TRANSLATED="&"E"&";"&"CURRENCY="&' +
             currency + '&";"&"BOOK="&' + # Currency
             book + '&";",' +             # Book
             year + ',' +                 # Year
             per_ltd + ',' +              # PER or LTD
             str(month) + ',' +           # Period
             div + ',"*","*",' +          # Div
             acct + ',"*","*")')          # GL Acct

# Determines correct Hierarchy for each GL
def scan_row(r):
    # Determines Report Type
    if r[2].strip() in ['Total assets',
                        'Total Liabilities',
                        'Total Equity']:
        cat['report'] = 'BS'
    elif r[2].strip() in ['Net sales', 'Cost of Sales',
                          'Operating expenses',
                          'Other expenses (income)',
                          'Current income tax expense']:
        cat['report'] = 'IS'

    # Determines Category A, B, C
    if r[0] == '' and r[1].isdigit() and int(r[1]) <= 10:
        cat['a'] = ' ' .join([str(int(r[1])), r[2]])
    elif r[1] == '' and r[2].isdigit() and len(str(int(r[2]))) == 2:
        cat['b'] = ' ' .join([str(int(r[2])), r[3]])
    elif r[2] == '' and r[3].isdigit() and len(str(int(r[3]))) == 3:
        cat['c'] = ' ' .join([str(int(r[3])), r[4]])

    # Fixes Other exoenses (income) / Current income tax Expense Hierarchy
    if cat.get('a'):
        if ('Other expenses (income)' in cat['a']
        or 'Current income tax expense' in cat['a']):
            if r[1] == '' and r[2].isdigit() and len(str(int(r[2]))) == 3:
                cat['b'] = ' ' .join([str(int(r[2])), r[3]])
                cat['c'] = ' ' .join([str(int(r[2])), r[3]])

    for cell in r:
        # 6 Digit Identification
        if cell.isdigit():
            if len(str(int(cell))) == 6:
                glacct = cell
                desc = r[r.index(cell)+1]
                gl[glacct] = (desc, cat['report'], cat['a'],
                              cat['b'], cat['c'])

############### utils ###############

def extract_gl():
    with open(input1) as i1:
        i1r = csv.reader(i1)
        for r in i1r:
            scan_row(r)
    print 'extract_gl() complete.'

def generate_rows():
    for acct in gl:
        desc, report, a, b, c = gl[acct]
        for year, div in product(years, divs):

            # Division / Currency / Desc
            currency = div_book[div][0]
            book = div_book[div][1]
            div_desc = div_book[div][2]

            if report == 'IS':
                #Writes X 12 for # of Months
                for month in range(1,13):
                    qtr = quarterperiod[month]
                    minus = '-'
                    per_ltd = '"PER"'
                    formula = makeformula(minus, currency, book, year,
                                          per_ltd, str(month), div, acct)
                    row = (acct, desc, report, a, b, c, div_desc, month,
                           qtr, year, str(formula))
                    rows.append(row)

            elif report == 'BS':
                # Writes X4 for # of Quarters
                for month in [3, 6, 9, 12]:
                    qtr = quarterperiod[month]
                    minus = ''
                    per_ltd = '"LTD"'
                    formula = makeformula(minus, currency, book, year,
                                          per_ltd, str(month), div, acct)
                    row = (acct, desc, report, a, b, c, div_desc, month,
                           qtr, year, str(formula))
                    rows.append(row)
    print 'generate_rows() complete.'

def write_csv():
    headers = ['GL', 'GL Desc', 'Report Type', 'Cat A', 'Cat B', 'Cat C',
               'Div', 'Period', 'Quarter', 'Year', 'Amount']

    with open(output, 'wb') as o1:
        o1w = csv.writer(o1)
        o1w.writerow(headers)
    
        for row in rows:
            o1w.writerow(row)
    print 'write_csv() complete.'

############## oracle_main() ###############

def oracle_main():
    t0 = time.clock()
    extract_gl()
    generate_rows()
    write_csv()
    t1 = time.clock()
    print 'oracle_main() completed. Duration:', t1-t0

if __name__ == '__main__':
    oracle_main()