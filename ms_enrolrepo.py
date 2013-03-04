import csv
import time
from aux_reader import *
from decimal import Decimal
import dateutil.parser as dparser
from collections import defaultdict
from datetime import datetime, timedelta

output = 'o\\Enrol Repo - 2013-03-04.csv'
input1 = 'i\\enrol_repo\\future billing - 2013-03-01.csv'
input2 = 'i\\enrol_repo\\contract repo - 2013-03-01.csv'

#################################################################################
## Function Definitions

# Declare Variables
def fb_datascrape(r):
    an = r['Agreement Number']
    if r['Imputed Rev'] == '': r['Imputed Rev'] = 0
    if r['GP'] == '': r['GP'] = 0
    ir = Decimal(r['Imputed Rev'])
    gp = Decimal(r['GP'])
    asd = dparser.parse(r['Agreement Start Date']).date()
    aed = dparser.parse(r['Agreement End Date']).date()
    sbd = dparser.parse(r['Scheduled Bill Date']).date()
    region = r['Region']
    level = r['Level']
    div = r['Div']
    program = r['Program Offering Code']
    cata = r['Custom Category A']
    catb = r['Custom Category B']
    catd = r['Custom Category D']
    adc = r['Agreement Desktop Count']
    cust = r['Primary Customer Name']
    if r['Agreement Status'] == 'Active':
        renewal = 'B'
    elif r['Agreement Status'] == 'Renewal Assumption - Annual Bill * 90%':
        renewal = 'R'
        ir = ir/Decimal(0.9)
        gp = gp/Decimal(0.9)
    ankey = an + ' ' + sbd.strftime("%Y-%m")

    # Time to scrape!
    if catd in monthlyfees:
        mgp = gp
        tgp = Decimal(0)
    else:
        mgp = Decimal(0)
        tgp = gp

    if ankey not in enroldict:
        try:
            enroldict[ankey] = [an, ir, tgp, mgp, asd, aed, sbd, region,
                                level, div, program, cata, catb, adc, cust,
                                renewal]
        except:
            print ankey
            raw_input('Error Push any key to continue...')   
    else:
        enroldict[ankey][1] += ir
        enroldict[ankey][2] += tgp
        enroldict[ankey][3] += mgp
    return

def cleanr(oldr, r):
    r['Divloc Des'] = oldr['Master Divloc Des']
    r['Sales Rep'] = oldr['Master OB Rep Name']
    r['SCC Master #'] = oldr['Master Number']
    r['Status'] = oldr['Contract Status']

    # Future Billing Data Parse
    if oldr['Contract Number'] in revenroldict:
        enr = revenroldict[oldr['Contract Number']]
        r['Enrol #'] = enr[0]
        r['Enrol Start Date'] = enr[4]
        r['Enrol End Date'] = enr[5]
        r['EA Program Type'] = oldr['Contract Program Name']
        r['Level'] = enr[8]
        r['Desktop Count'] = enr[13]
        r['Primary Customer Name'] = enr[14]
        r['Div'] = enr[9]
        r['Region'] = enr[7]
        r['Estimated Annual Rev'] = enr[1]
        r['Period'] = enr[6]
        r['Estimated Trans GP'] = enr[2]
        r['Estimated Monthly GP'] = enr[3]
        r['2013 Rev Stream'] = enr[15]
        if enr[11] == 'ESA 2.0':
            r['ESA 2.0 / 3.0 / NON-EA'] = 'ESA 2.0'
        elif enr[11] == 'NON-EA DIRECT':
            r['ESA 2.0 / 3.0 / NON-EA'] = 'NON-EA DIRECT'
        else:
            r['ESA 2.0 / 3.0 / NON-EA'] = 'ESA 3.0'
            r['New / Renewal'] = enr[12]
            if enr[11] == 'ESA 3.0 - CORPORATE':
                r['Corporate / Major'] = 'Corporate'
            else:
                r['Corporate / Major'] = 'Major'

    # Not from Future billing, Data Parse
    else:
        r['Enrol #'] = oldr['Contract Number']
        r['Enrol Start Date'] = oldr['Contract Start Date']
        r['Enrol End Date'] = oldr['Contract End Date']
        r['EA Program Type'] = oldr['Contract Program Name']
        r['Level'] = oldr['Contract Level']
        r['Desktop Count'] = oldr['Contract Units']
        r['Primary Customer Name'] = oldr['Master Name']
        r['Div'] = oldr['Master Division']
        r['Estimated Annual Rev'] = None
        r['Period'] = None
        r['Estimated Trans GP'] = None
        r['Estimated Monthly GP'] = None
        r['ESA 2.0 / 3.0 / NON-EA'] = None
        r['New / Renewal'] = oldr['Contract Category']
        r['Corporate / Major'] = None
        divloc = oldr['Master Division'] + oldr['Master Divloc']
        if divloc in divregion:
            r['Region'] = divregion[divloc]
    return r

# Pictionary Jars
enroldict = defaultdict(str)
revenroldict = defaultdict(str)
monthlyfees = ['Deploy Fee', 'Help Desk Fee', 'Manage Fee']
divregion = csv_dic('auxiliary\\div-region.csv')

#################################################################################
## Main

def main():
    t0 = time.clock()

    header = ('Enrol #', 'Enrol Start Date', 'Enrol End Date',
              'Corporate / Major', 'ESA 2.0 / 3.0 / NON-EA', 'EA Program Type',
              'New / Renewal', 'Level','Desktop Count', 'SCC Master #',
              'Primary Customer Name', 'Div', 'Region', 'Divloc Des', 'Sales Rep',
              'Period','Estimated Annual Rev', 'Estimated Trans GP',
              'Estimated Monthly GP', 'Status', '2013 Rev Stream')

    # Future Billing Data Scrape
    with open(input1) as i1:
        i1r = csv.DictReader(i1)
        for r in i1r: fb_datascrape(r)

    # Revised EnrolDict
    for enrol in enroldict:
        if enroldict[enrol][1] > Decimal(0):
            revenroldict[enroldict[enrol][0]] = enroldict[enrol]

    # Output Writer
    with open(output, 'wb') as o:
        writer = csv.writer(o)
        ow = csv.DictWriter(o, fieldnames=header)
        writer.writerows([header])

        # Contract Repository
        with open(input2) as i2:
            i2r = csv.DictReader(i2)
            for oldr in i2r:
                r = dict.fromkeys(header)
                r = cleanr(oldr, r)
                ow.writerow(r)
    t1 = time.clock()
    return 'Process completed! Duration:', t1-t0

print main()
