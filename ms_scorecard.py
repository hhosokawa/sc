import csv
import time
from decimal import Decimal
from datetime import datetime, timedelta
import dateutil.parser as dparser
from aux_reader import *

output = 'o\\04-Jan-13 Enrol Repo.csv'
input1 = 'i\\ms_scorecard\\referrals.csv'
currentyear = datetime(2013, 1, 1).date()
currentyearend = datetime(2014, 1, 1).date()

growth = tree()
renewal = tree()
trueup = tree()
divregion = csv_dic('auxiliary\\div-region.csv')
divdistrict = csv_dic('auxiliary\\div-district.csv')
custclass = csv_dic('auxiliary\\cust-class.csv')

#################################################################################
## Function Definitions

# EA Growth - Datascrape: Referrals -> EA Growth Data
def growth_scrape(r):
    enrol = r['Contract Enrollment']
    startdate = dparser.parse(r['Enrollment Start Date']).date()
    if (enrol not in growth and startdate >= currentyear 
    and r['Agreement Category'] == 'New Contract'):
        growth[enrol]['Cust Name'] = r['SoftChoice Customer Name']
        growth[enrol]['Cust Class'] = custclass.get(r['SoftChoice Customer Number'], '')
        growth[enrol]['Start Date'] = r['Enrollment Start Date']
        growth[enrol]['Start Month'] = int(startdate.strftime("%m"))
        divloc = r['Referral Division'] + r['Referral Divloc Number']
        growth[enrol]['Region'] = divregion.get(divloc, '')
        growth[enrol]['District'] = divdistrict.get(divloc, '')        
        if (r['Contract Type'] 
        in ['MS ECI Enrollment', 'MS EAP Enrollment', 'MS Enterprise Enrollment']):
            growth[enrol]['Contract Type'] = r['Contract Type']
        else:
            growth[enrol]['Contract Type'] = 'Other'
    return

# Renewals - Datascrape: Referrals -> Renewal Execution
def renewal_scrape(r):
    enrol = r['Contract Enrollment']
    startdate = dparser.parse(r['Enrollment Start Date']).date()
    enddate = dparser.parse(r['Enrollment End Date']).date()
    if (enrol not in renewal) and (currentyear <= enddate <= currentyearend):
        renewal[enrol]['Start Date'] = startdate
        renewal[enrol]['Referral #'] = r['Referral Number']
    elif (enrol in renewal
    and renewal[enrol]['Start Date'] <> startdate
    and renewal[enrol]['Referral #'] <> r['Referral Number']):
        renewal[enrol]['Renew'] = True
        print enrol, renewal[enrol]
        raw_input('...')
    return


#################################################################################
## Main

def main():
    t0 = time.clock()

    # Create Headers
    header = ('Enrol #', 'Period', 'Contract Type', 'Cust Name', 
              'Cust Class', 'Region', 'District')

    # Referral Datascrape
    with open(input1) as i1:
        i1r = csv.DictReader(i1)
        for r in i1r:
            growth_scrape(r)
            renewal_scrape(r)
        raw_input('... Too far')

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
