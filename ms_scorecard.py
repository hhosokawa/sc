import csv
import time
from decimal import Decimal
from datetime import datetime, timedelta
import dateutil.parser as dparser
from aux_reader import *

output = 'o\\04-Jan-13 Enrol Repo.csv'
input1 = 'i\\ms_scorecard\\referrals.csv'
input2 = 'i\\ms_scorecard\\contract repo.csv'
currentyear = datetime(2013, 1, 1).date()
currentyearend = datetime(2014, 1, 1).date()

d = Decimal
growth = tree()
renew = tree()
enrolgp = tree()
divregion = csv_dic('auxiliary\\div-region.csv')
divdistrict = csv_dic('auxiliary\\div-district.csv')
custclass = csv_dic('auxiliary\\cust-class.csv')

#################################################################################
## Function Definitions

# Determines if row is an annual bill
def annualbill(r):
    if ((r['Product Item Bill Type'] in ['OAP2', 'OAP3']) or
        (r['Product Item Number'] in ['KM6165', 'KM6166'])): return True
    return False

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

    # Absorb Enrol:GP Dict
    if annualbill(r) and d(r['Referral Net Expected Fee Total']) > 0:
        gp = d(r['Referral Net Expected Fee Total']) + d(r['Referral Partner Fee'])
        enrolgp[enrol] = gp
    return

# Renewal Analysis - Datascrape: Contract Repo -> Renewal Analysis
def renew_scrape(r):
    enrol = r['Contract Number']
    enddate = dparser.parse(r['Contract End Date']).date()
    if enrol not in renew and currentyear <= enddate <= currentyearend:
        renew[enrol]['End Date'] = enddate.strftime("%Y-%m")
        renew[enrol]['Program Name'] = r['Contract Program Name']
        renew[enrol]['Level'] = r['Contract Level']
        divloc = r['Master Division'] + r['Master Divloc']
        renew[enrol]['Region'] = divregion.get(divloc, '')
        renew[enrol]['District'] = divdistrict.get(divloc, '')
        renewid = r['Renewed to Vendor Contract ID']
        if renewid == '': 
            renew[enrol]['Renew'] = ''
        else: 
            renew[enrol]['Renew'] = 'True'
    return
        
#################################################################################
## Main

def main():
    t0 = time.clock()

    # Create Headers
    header = ('Enrol #', 'Period', 'Contract Type', 'Cust Name', 
              'Cust Class', 'Region', 'District')

    # Referral Datascrape -> EA Growth Analysis
    with open(input1) as i1:
        for r in csv.DictReader(i1): growth_scrape(r)

    print len(enrolgp)
    raw_input('...')

    # Contract Repo Datascrape -> Renewal Analysis
    with open(input2) as i2:
        for r in csv.DictReader(i2): renew_scrape(r)

    for x in renew:
        if renew[x]['Renew'] == 'True':  print x

    t1 = time.clock()
    return 'Process completed! Duration:', t1-t0

print main()

"""
    # Output Writer
    with open(output, 'wb') as o:
        writer = csv.writer(o)
        ow = csv.DictWriter(o, fieldnames=header)
        writer.writerows([header])
"""

