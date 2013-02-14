import csv
import time
from decimal import Decimal
from datetime import datetime, timedelta
import dateutil.parser as dparser
from aux_reader import *

output = 'o\\04-Jan-13 Enrol Repo.csv'
input1 = 'i\\ms_scorecard\\referrals.csv'
input2 = 'i\\ms_scorecard\\contract repo.csv'
currentyear = datetime(2012, 1, 1).date()
currentyearend = datetime(2013, 1, 1).date()

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
    if ((r['Product Item Bill Type'] in ['OAP2', 'OAP3']) 
    or (r['Product Item Number'] in ['KM6165', 'KM6166'])): 
        return True
    else:
        return False

# EA Growth - Datascrape: Referrals -> EA Growth Data
def growth_scrape(r):
    enrol = r['Contract Enrollment']
    startdate = dparser.parse(r['Enrollment Start Date']).date()
    enddate = dparser.parse(r['Enrollment End Date']).date()
    if (enrol not in growth and startdate >= currentyear 
    and r['Agreement Category'] == 'New Contract'):
        x = growth[enrol] # id reference
        x['Enrol #'] = enrol
        x['Cust Name'] = r['SoftChoice Customer Name']
        x['Cust Class'] = custclass.get(r['SoftChoice Customer Number'], '')
        x['Start Date'] = startdate.strftime("%Y-%m")
        x['Period'] = int(startdate.strftime("%m"))
        x['End Date'] = enddate.strftime("%Y-%m")
        divloc = r['Referral Division'] + r['Referral Divloc Number']
        x['Region'] = divregion.get(divloc, '')
        x['District'] = divdistrict.get(divloc, '')
        x['Report Type'] = 'Growth EA'
        if (r['Contract Type'] 
        in ['MS ECI Enrollment', 'MS EAP Enrollment', 'MS Enterprise Enrollment']):
            x['Custom Category A'] = r['Contract Type']
        else:
            x['Custom Category A'] = 'Other'

    # Absorb Enrol:GP Dict
    if annualbill(r) and d(r['Referral Net Expected Fee Total']) > 0:
        gp = d(r['Referral Net Expected Fee Total']) + d(r['Referral Partner Fee'])
        enrolgp[enrol] = gp
    return

# Renewal Analysis - Datascrape: Contract Repo -> Renewal Analysis
def renew_scrape(r):
    enrol = r['Contract Number']
    startdate = dparser.parse(r['Contract Start Date']).date()
    enddate = dparser.parse(r['Contract End Date']).date()
    if enrol not in renew and currentyear <= enddate < currentyearend:
        x = renew[enrol] # id reference
        x['Enrol #'] = enrol
        x['Cust Name'] = r['Master Name']
        x['Cust Class'] = r['Contract Level']
        x['Start Date'] = startdate.strftime("%Y-%m")
        x['End Date'] = enddate.strftime("%Y-%m")
        x['Period'] = int(enddate.strftime("%m"))
        divloc = r['Master Division'] + r['Master Divloc']
        x['GP'] = enrolgp.get(r['Renewed to Vendor Contract ID'], 0)
        x['Region'] = divregion.get(divloc, '')
        x['District'] = divdistrict.get(divloc, '')
        x['Report Type'] = 'Renewal Analysis'
        renewid = r['Renewed to Vendor Contract ID']
        if renewid == '': 
            x['Custom Category A'] = 'Not Renewed'
        else: 
            x['Custom Category A'] = 'Renewed'
    return
        
#################################################################################
## Main

def main():
    t0 = time.clock()

    # Create Headers
    header = ('Enrol #', 'Cust Name', 'Cust Class', 'Region', 
              'District', 'Start Date', 'End Date', 'Period', 
              'Level', 'GP', 'Custom Category A', 'Report Type')

    # Referral Datascrape -> EA Growth Analysis
    with open(input1) as i1:
        for r in csv.DictReader(i1): growth_scrape(r)
        print 'EA Growth Analysis Complete.', time.clock()-t0

    # Contract Repo Datascrape -> Renewal Analysis
    with open(input2) as i2:
        for r in csv.DictReader(i2): renew_scrape(r)
        print 'Renewal Analysis Complete.', time.clock()-t0

    # Write to Output
    with open(output, 'wb') as o:
        writer = csv.writer(o)
        ow = csv.DictWriter(o, fieldnames=header)
        writer.writerows([header])
        for enr in growth: ow.writerow(growth[enr])
        for enr in renew: ow.writerow(renew[enr])

    t1 = time.clock()
    return 'Process completed! Duration:', t1-t0

print main()
