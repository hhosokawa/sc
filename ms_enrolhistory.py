import csv
import time
import dateutil.parser as dparser
from datetime import datetime, timedelta
from collections import defaultdict
from decimal import Decimal

output = 'o\\10-Dec-12 - enrol history.csv'
input1 = 'i\\2010-2012 - Enrollment PAST.csv'
refset = set()
enroldict = defaultdict(str)

#################################################################################
## Function Definitions

def annualbill(r):
    if ((r['Product Item Bill Type'] in ['BDEA', 'OAP2', 'OAP3']) or
        (r['Product Item Number'] in ['KM6165', 'KM6166'])): return True
    return False

def enrol_analysis(r):
    ce = r['Contract Enrollment']
    pibt = r['Product Item Bill Type']
    pin = r['Product Item Number']
    rn = r['Referral Number']
    rs = r['Referral Source']
    try: rts = Decimal(r['Referral Total Sum'])
    except: rts = 0
    refdate = dparser.parse(r['Referral Date (DD-MMM-YYYY)']).date()
# [Latest Annual Bill Date, Annual Billing $, True-Up $, Add-On $, Ref Source, Bill Type]
    if (annualbill(r)) and (ce in enroldict) and (rts > 0):
        if refdate > enroldict[ce][0]:
            enroldict[ce] = [refdate, rts, 0 ,0, rs, pibt]
            refset.add(rn)
    elif (annualbill(r)) and (rts > 0):
        enroldict[ce] = [refdate, rts, 0 ,0, rs, pibt]
        refset.add(rn)
    elif (pibt == 'TRUP') and (ce in enroldict) and (rn not in refset):
        if refdate > enroldict[ce][0]:
            enroldict[ce][2] += rts
            refset.add(rn)
    elif (pibt == 'ADOT') and (ce in enroldict) and (rn not in refset):
        if refdate > enroldict[ce][0]:
            enroldict[ce][3] += rts
            refset.add(rn)
    return

#################################################################################
## Main Program

def main():
    t0 = time.clock()
    with open(input1) as i1:
        i1r = csv.DictReader(i1)
        for r in i1r: enrol_analysis(r)
    with open(output, 'wb') as o1:
        wr = csv.writer(o1)
        wr.writerows([('Enrollment_#', 'Date', 'Annual_Bill_$',
                       'True_Up', 'Add_On', 'Ref Source', 'Bill Type')])
        for k,v in enroldict.items():
            new_v = map(str, v)
            billdate, price, trueup, addon, refsource, billtype = new_v
            wr.writerow([k, billdate, price, trueup, addon, refsource, billtype])
    t1 = time.clock()
    return 'Process completed! Duration:', t1-t0

print main()
