import csv
import time
from collections import defaultdict
from decimal import *

output1 = 'o\\2010-2013 Devan Freq 26-Nov-12 - 1.csv'
output2 = 'o\\2010-2013 Devan Freq 26-Nov-12 - 2.csv'
input1 = 'i\\dmoo - 2010 so gp freq.csv'
input2 = 'i\\dmoo - 2011 so gp freq.csv'
input3 = 'i\\dmoo - 2012 so gp freq.csv'
input4 = 'i\\dmoo - 2013 so gp freq.csv'
input5 = 'i\\dmoo - 2010-2013 ref gp freq.csv'

# this is a change

#################################################################################
## Function Definitions

def header_add(header):             # Adds New Headers
    newfields = ['Super Cat', 'Quarter', 'Region', 'SO #',
                 'SCC Cat', 'GP', 'Year', 'Ref / Sales']
    for newfield in newfields: header.add(newfield)
    return header

def csv_dic(filename):              # Converts CSV to dict
    reader = csv.reader(open(filename, "rb"))
    my_dict = dict((k, v) for k, v in reader)
    return my_dict

def region(div, divloc):            # Region Adjustment
    key = div + divloc
    try:
        return divregion[key]
    except:
        return 'N/A'

def datascrape(r):
    key = r['Order Number'] + ' - ' + r['Item SCC Category']
    if key not in invdict:
        r['Total GP $'] = Decimal(r['Total GP $'])
        r['Item Number'] = 'Sales'
        invdict[key] = r
    else:
        invdict[key]['Total GP $'] += Decimal(r['Total GP $'])
        invdict[key]['Total GP $'].quantize(Decimal('.01'), rounding=ROUND_UP)
    return

def refdatascrape(r):
    if r['Referral Number'] not in invdict:
        r['Referral Revenue Total'] = Decimal(r['Referral Revenue Total'])
        r['Product Item Desc'] = 'Referral'
        invdict[r['Referral Number']] = r
    else:
        pass
    return

def mapheader(oldr, r, ref_sales):
    if ref_sales == 'Sales':
        r['Ref / Sales'] = 'Sales'
        r['Super Cat'] = oldr['Item Super Category']
        r['SCC Cat'] = oldr['Item SCC Category']
        r['Quarter'] = oldr['Invoice date (SC FQ)']
        r['SO #'] = oldr['Order Number']
        r['GP'] = oldr['Total GP $']
        r['Year'] = Decimal(oldr['Invoice date (SC FY)']) - 1
        r['Region'] = region(oldr['Softchoice Division'],
                             oldr['Softchoice DivLoc'])
    elif ref_sales == 'Referral':

        r['Ref / Sales'] = 'Referral'
        r['Super Cat'] = oldr['Referral Super Category']
        r['Quarter'] = 'Q' + oldr['Referral Fiscal Quarter']
        r['SO #'] = oldr['Referral Number']
        r['GP'] = Decimal(oldr['Referral Revenue Total'])
        r['Year'] = Decimal(oldr['Referral Fiscal Year']) - 1
        r['Region'] = region(oldr['Referral Division'],
                             oldr['Referral Divloc Number'])
    return r

#################################################################################
##  Dictionary Pictionary Jars

divregion = csv_dic('auxiliary\\div-region.csv')
invdict = defaultdict(str)
tempdict = defaultdict(str)

#################################################################################
## Main Program

def main():
    t0 = time.clock()
    header = set()                          # Obtain Correct Headers
    header = header_add(header)
    header = tuple(header)

    with open(input1) as i1:                # Sales Input 1
        i1r = csv.DictReader(i1)
        for r in i1r: datascrape(r)
    with open(input2) as i2:                # Sales Input 2
        i2r = csv.DictReader(i2)
        for r in i2r: datascrape(r)
    with open(output1, 'wb') as o:          # Output Writer
        writer = csv.writer(o)
        ow = csv.DictWriter(o, fieldnames=header)
        writer.writerows([header])
        for i in invdict:
            r = dict.fromkeys(header)
            r = mapheader(invdict[i], r, 'Sales')
            ow.writerow(r)
    invdict.clear()

    with open(input3) as i3:                # Sales Input 3
        i3r = csv.DictReader(i3)
        for r in i3r: datascrape(r)
    with open(input4) as i4:                # Sales Input 4
        i4r = csv.DictReader(i4)
        for r in i4r: datascrape(r)
    with open(output2, 'wb') as o:          # Output Writer
        writer = csv.writer(o)
        ow = csv.DictWriter(o, fieldnames=header)
        writer.writerows([header])
        for i in invdict:
            r = dict.fromkeys(header)
            r = mapheader(invdict[i], r, 'Sales')
            ow.writerow(r)
        invdict.clear()
        with open(input5) as i5:            # Referral Input 5
            i5r = csv.DictReader(i5)
            for r in i5r: refdatascrape(r)
            for i in invdict:
                r = dict.fromkeys(header)
                r = mapheader(invdict[i], r, 'Referral')
                ow.writerow(r)
    invdict.clear()
    t1 = time.clock()
    return 'Process completed! Duration:', t1-t0

print main()
