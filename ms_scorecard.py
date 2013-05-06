import csv
import time
from decimal import Decimal as d
from datetime import datetime, timedelta
import dateutil.parser as dparser
from aux_reader import *

output = 'o\\MS Scorecard - 2013-04-02.csv'
input1 = 'i\\ms_scorecard\\referrals.csv'
input2 = 'i\\ms_scorecard\\contract repo.csv'
input3 = 'i\\ms_scorecard\\explore - zero.csv'
custclass = csv_dic('i\\ms_scorecard\\cust-class.csv')
currentyear = datetime(2013, 1, 1).date()
currentyearend = datetime(2014, 1, 1).date()

growth = tree()
renew = tree()
trueup = tree()
enroltu = set()
zerotu = set()
enrolgp = tree()
divregion = csv_dic('auxiliary\\div-region.csv')
divdistrict = csv_dic('auxiliary\\div-district.csv')

#################################################################################
## Function Definitions

# Determines if row is an annual bill
def annualbill(r):
    if ((r['Product Item Bill Type'] in ['OAP2', 'OAP3'])
    or (r['Product Item Number'] in ['KM6165', 'KM6166'])):
        return True
    else:
        return False

# Determines if row is a True-Up
def trueupcheck(r):
    if (r['Referral Rev Type'] == 'TU'
    or r['Product Item Bill Type'] == 'TRUP'):
        return True
    else:
        return False

# Datascrape: Referrals -> EA Growth Data
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

    # Absorb Enrol:True-Up Dict
    yearprior = int(r['Referral Fiscal Year'])-2
    if (trueupcheck(r) and yearprior == int(currentyear.strftime("%Y"))-1):
        enroltu.add(enrol)
    return

# Remap Contract Repo -> Dict Headers
def remap(r, style):
    enrol = r['Contract Number']
    startdate = dparser.parse(r['Contract Start Date']).date()
    enddate = dparser.parse(r['Contract End Date']).date()

    # Identify Renewal / True-Up Executed
    if style == 'renew':
        x = renew[enrol]
        reporttype = 'Renewal Analysis'
        x['Period'] = int(enddate.strftime("%m"))
        if r['Renewed to Vendor Contract ID'] <> '':
            x['Custom Category A'] = 'A) Renewed'
            x['GP'] = enrolgp.get(r['Renewed to Vendor Contract ID'], 0)
        else:
            x['Custom Category A'] = 'B) Not Renewed'
    elif style == 'trueup':
        x = trueup[enrol]
        reporttype = 'True-Up Analysis'
        x['Period'] = int(startdate.strftime("%m"))
        if (enrol in enroltu) or (enrol in zerotu):
            x['Custom Category A'] = 'A) Executed'
        else:
            x['Custom Category A'] = 'B) Not Executed'

    x['Enrol #'] = enrol
    x['Cust Name'] = r['Master Name']
    x['Cust Class'] = r['Contract Level']
    x['Start Date'] = startdate.strftime("%Y-%m")
    x['End Date'] = enddate.strftime("%Y-%m")
    divloc = r['Master Division'] + r['Master Divloc']
    x['Region'] = divregion.get(divloc, '')
    x['District'] = divdistrict.get(divloc, '')
    x['Report Type'] = reporttype
    return

# Datascrape: Contract Repo -> Renewal/True-Up Analysis
def renew_scrape(r):
    enrol = r['Contract Number']
    enddate = dparser.parse(r['Contract End Date']).date()

    if (enrol not in renew
    and currentyear <= enddate < currentyearend):
        remap(r, 'renew')
    if (enrol not in trueup
    and enddate >= currentyearend):
        remap(r, 'trueup')
    return

# Datascrape: Explore.ms Purchase Order - Zero Usage -> Set()
def zero_scrape(r):
    if r['Agreement Number'] not in zerotu:
        zerotu.add(r['Agreement Number'])
    return

#################################################################################
## Main

def main():
    t0 = time.clock()

    # Create Headers
    header = ('Enrol #', 'Cust Name', 'Cust Class', 'Region',
              'District', 'Start Date', 'End Date', 'Period',
              'Level', 'GP', 'Custom Category A', 'Report Type')

    # Zero True-Up Transactions -> Set()
    with open(input3) as i3:
        for r in csv.DictReader(i3): zero_scrape(r)
        print 'Zero True-up Anaylsis Complete.', time.clock()-t0

    # Referral Datascrape -> EA Growth Analysis
    with open(input1) as i1:
        for r in csv.DictReader(i1): growth_scrape(r)
        print 'EA Growth Analysis Complete.', time.clock()-t0

    # Contract Repo Datascrape -> Renewal /True Up Analysis
    with open(input2) as i2:
        for r in csv.DictReader(i2): renew_scrape(r)
        print 'Renewal / True-Up Analysis Complete.', time.clock()-t0

    # Write to Output
    with open(output, 'wb') as o:
        writer = csv.writer(o)
        ow = csv.DictWriter(o, fieldnames=header)
        writer.writerows([header])
        for enr in growth: ow.writerow(growth[enr])
        for enr in renew: ow.writerow(renew[enr])
        for enr in trueup: ow.writerow(trueup[enr])
    t1 = time.clock()
    return 'Process completed! Duration:', t1-t0

print main()
