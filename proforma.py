import csv
import time
from aux_reader import *

#################################################################################
## Function Definitions
output = 'o\\proforma.csv'
input1 = 'i\\proforma\\gl.csv'

years = ['2010', '2011', '2012']
div = '"*"'
quarterperiod = {1: 'Q1', 2: 'Q1', 3: 'Q1', 
                 4: 'Q2', 5: 'Q2', 6: 'Q2', 
                 7: 'Q3', 8: 'Q3', 9: 'Q3', 
                 10: 'Q4', 11: 'Q4', 12: 'Q4'}

gl = tree()
cat = tree()

#################################################################################
## Function Definitions

# Obtain GLs
def scan_gl(r):

    # Determines Report Type
    if r[2] in ['Assets', 'Liabilities', 'Equity']:
        cat['report'] = 'BS'
    elif r[2] in ['Revenue', 'Cost of Sales', 'Expenses']:
        cat['report'] = 'IS'

    # Determines Category A,  B
    if r[1] != '' and r[1].isdigit() and len(str(r[1])) == 5:
        cat['a'] = ' ' .join([r[1], r[2]])
    elif r[2] != '' and r[2].isdigit() and len(str(r[2])) == 5:
        cat['b'] = r[3]

    # Determines Category C
    if r[3] != '' and r[3].isdigit() and len(str(r[3])) == 5:
        cat['c'] = r[4]

    for cell in r:
        # 6 Digit Identification
        if len(str(cell)) == 6 and cell.isdigit():
            glacct = cell
            desc = r[r.index(cell)+1] 
            if cat['report'] == 'IS': # IS
                gl[glacct] = (desc, cat['report'], cat['a'], cat['b'], desc)
            else: # BS
                gl[glacct] = (desc, cat['report'], cat['a'], cat['b'], cat['c'])
    return

# GL Account -> Write Row for Year, IS, BS
def mapper(o, acct, desc, report, a, b, c):

    # Cycles # Of Years
    for year in years:

        if report == 'IS': # IS
            # Writes X 12 for # of Months
            for month in range(1,13):
                qtr = quarterperiod[month]
                formula = ('=-GXL("A","","TRANSLATED="&"E"&";"&"CURRENCY="&"USD"&";"&"BOOK="&"SC Consol - US"&";",' + 
                           year + ',"PER",' +   # Year
                           str(month) + ',' +   # Period
                           div + ',"*","*",' +  # Div
                           acct + ',"*","*")')  # GL Acct

                o.writerow([acct, desc, report, a, b, c, div, month, qtr, year, str(formula)])

        else: # BS
            # Writes X 4 for # of Quarters
            for month in [3, 6, 9, 12]:
                qtr = quarterperiod[month]
                formula = ('=GXL("A","","TRANSLATED="&"E"&";"&"CURRENCY="&"USD"&";"&"BOOK="&"SC Consol - US"&";",' + 
                           year + ',"LTD",' +   # Year 
                           str(month) + ',' +   # Period / Quarter
                           div + ',"*","*",' +  # Div
                           acct + ',"*","*")')  # GL Acct

                # "Due to Parent Company" ->  "Long Term Assets"
                if c == 'Due to Parent Company':
                    a = '10000 Assets'
                    b = 'Long Term Assets'

                o.writerow([acct, desc, report, a, b, c, div, month, qtr, year, str(formula)])
                  
    return


#################################################################################
## Main Program

def main():
    t0 = time.clock()

    # Collect all Valid GLs
    with open(input1) as i1:
        i1r = csv.reader(i1)
        for r in i1r: r = scan_gl(r)

    headers = ['GL', 'GL Desc', 'Report Type', 'Cat A', 'Cat B', 'Cat C', 
               'Div', 'Period', 'Quarter', 'Year', 'Amount']

    with open(output, 'wb') as o1:
        o1w = csv.writer(o1)
        o1w.writerow(headers)
        for acct in gl:
            desc, report, a, b, c = gl[acct]
            mapper(o1w, acct, desc, report, a, b, c)

    t1 = time.clock()
    return 'Process completed! Duration:', t1-t0

print main()
