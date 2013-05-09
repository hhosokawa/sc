import csv
import time
from aux_reader import *

#################################################################################
## Function Definitions
output = 'o\\VIOZ Recurring - 2011-2013.csv'
input1 = 'i\\vioz_recurring\\2011-2013 - ref.csv'
input2 = 'i\\vioz_recurring\\2012-2013 - sales.csv'
input3 = 'i\\vioz_recurring\\2011 - sales.csv'
acct_master = csv_dic('i\\vioz_recurring\\acct - master.csv')

new_headers = ['master class', 'unique master', 'NAICS industry', 'NAICS desc',
               'product type', '']

master_naics = csv_dic('auxiliary\\master id - NAICS.csv', 6)
FX = csv_dic('auxiliary\\dictFX.csv')
majoraccts = csv_dic('auxiliary\\enrol - major customers.csv')
ref_agreementcat = csv_dic('auxiliary\\ref-agreementcategory.csv', '3b')

#################################################################################
## Dictionary Pictionary Jars

otherdiv = {'200':'100', '100':'200'}
divregion = csv_dic('auxiliary\\div-region.csv')
divdistrict = csv_dic('auxiliary\\div-district.csv')
venprogram = csv_dic('auxiliary\\dictvenprograms.csv')
ref_revtype = {'AO': 'EA Add-On',
               'N': 'EA New',
               'RC': 'EA Recurring',
               'RN': 'EA Renewal',
               'TU': 'EA True-Up'}

#################################################################################
## Function Definitions

# Adjusts for Finance's FX File
def adjFX(div, period, year):
    if div == 100: i = 'CDN-US-'
    else: i = 'US-CDN-'
    answer = i+str(period)+'-'+str(year)
    return float(FX[answer])

# Clean Referrals Data
def refclean(row):
    row['Quarter'] = 'Q' + str(row['Quarter'])
    row['Revenue Type'] = 'Referrals'
    row['Gross Revenue'] = row['GP']
    row['Custom Category A'] = row['Referral Source']
    row['Custom Category C'] = ''

    # ESA 3.0 Classification
    if row['Referral Source'] == 'ESA 3.0':

        # New Contracts VS Renewal Contracts
        try:
            if ref_agreementcat[row['Referral Number']] == 'New Contract':
                row['Custom Category B'] = 'New Contract'
            else:
                row['Custom Category B'] = 'Renew Contract'
        except KeyError:
            row['Custom Category B'] = 'New Contract'
        if row['Referral Rev Type'] == 'N':
            row['Custom Category B'] = 'New Contract'
        elif row['Referral Rev Type'] == 'RN':
            row['Custom Category B'] = 'Renew Contract'

        # Major VS Corporate
        if row['Referral Rev Type'] in ref_revtype:
            row['Custom Category C'] = ref_revtype[row['Referral Rev Type']]
        else:
            row['Custom Category C'] = 'Other'
        if (row['Customer Number'] in majoraccts
        or 'MAJOR' in row['Referral Notes']
        or row['Product Item Desc'] == 'HELP DESK INCENTIVE'):
            row['Custom Category A'] = 'ESA 3.0 - MAJOR'
        else: row['Custom Category A'] = 'ESA 3.0 - CORPORATE'

        # Correct Imputed Revenue for ESA 3.0
        if row['Referral Source'] == 'ESA 3.0':
            if (row['Product Item Desc'] == 'CORPORATE ACCOUNT' or
                row['Product Item Desc'] == 'MAJOR ACCOUNT' or
                row['Product Item Bill Type'] in ['ADOT', 'TRUP']): pass
            else: row['Imputed Revenue'] = 0

    # ESA 2.0 Classification
    elif row['Referral Source'] == 'ESA 2.0':
        if row['Product Item Bill Type'] == 'TRUP':
            row['Custom Category B'] = 'EA True Up'
        elif row['Product Item Bill Type'] == 'ADOT':
            row['Custom Category B'] = 'EA Add On'
        elif row['Product Item Bill Type'] == 'OAP2':
            row['Custom Category B'] = 'Year 2'
        elif row['Product Item Bill Type'] == 'OAP3':
            row['Custom Category B'] = 'Year 3'
        elif row['Product Item Bill Type'] == 'O':
            if ('Y1' or 'YR1') in row['Product Item Desc']:
                row['Custom Category B'] = 'Year 1'
            elif ('Y2' or 'YR2') in row['Product Item Desc']:
                row['Custom Category B'] = 'Year 2'
            elif ('Y3' or 'YR3') in row['Product Item Desc']:
                row['Custom Category B'] = 'Year 3'
            else:
                row['Custom Category B'] = 'Other'
        else:
            row['Custom Category B'] = 'Year 1'

    # MS SIP Classification
    elif row['Referral Source'] == 'MS SIP':
        row['Custom Category B'] = row['Product Item Desc'].title()

    # Other Classification
    else: row['Custom Category B'] = 'Other'

    # Ref PO 'CR' = 0 Rule
    if row['Referral PO Reference'] == 'CR':
        row['Imputed Revenue'] = 0
        row['Gross Revenue'] = 0
    return row

# Cleans Sales Data
def salesclean(row):
    row['Revenue Type'] = 'Sales'
    row['Imputed Revenue'] = row['Gross Revenue']
    row['Custom Category A'] = 'NON-EA DIRECT'
    row['Custom Category B'] = venprogram.get(row['Item Prodtype Venprogram'],
                                              'EA Indirect and Other')
    return row

# Region / District Adjustment
def region_district(row):
    divloc = str(row['Division']) + str(row['Divloc'])
    row['Region'] = divregion.get(divloc, 'N/A')
    row['District'] = divdistrict.get(divloc, 'N/A')
    return row


# Adds Product Type and Customer Type Details
def colorizer(r):

    # Rev Type - Recurring, Project, Run-Rate, Adjustments
    if r['Revenue Type'] == 'Referrals':
        r['product type'] = 'Recurring'
    else: # Need to Analyze SALES data
        if r['Item Revenue Recognition'] in ['HWMTN', 'SWMTN', 'SWSUB', 'SWBNDL']:
            r['product type'] = 'Recurring'
        else:
            try: grossrev = float(r['Gross Revenue'])
            except ValueError: grossrev = 0
            try: gp = float(r['GP'])
            except ValueError: gp = 0
            if (-1000 <= gp <= 1000) and (-10000 <= grossrev <= 10000):
                r['product type'] = 'Run Rate'
            else:
                r['product type'] = 'Project'

    # Cust Type - Master NAICS
    masterid = acct_master.get(r['Customer Number'], r['Customer Number'])
    try:
        r['master class'] = master_naics[masterid][1]
        r['unique master'] = master_naics[masterid][5]
        r['NAICS industry']= master_naics[masterid][2]
        r['NAICS desc']= master_naics[masterid][3]
    except KeyError:
        r['master class'] = 'N/A'
        r['unique master'] = 'N/A'
        r['NAICS industry']= 'N/A'
        r['NAICS desc']= 'N/A'

    return r


#################################################################################
## Main Program

def main():
    t0 = time.clock()
    header = set()

    # Get all input headers -> output header
    with open(input1) as i1: header.update(csv.DictReader(i1).fieldnames)
    with open(input2) as i2: header.update(csv.DictReader(i2).fieldnames)
    for x in new_headers: header.add(x)
    header = tuple(header)

    # Analyze Input 1-2 -> Output
    with open(output, 'wb') as o:
        writer = csv.writer(o)
        writer.writerows([header])
        with open(input1) as i1:
            i1r = csv.DictReader(i1)
            ow = csv.DictWriter(o, fieldnames=header)
            for row in i1r:
                row = refclean(row)
                row = region_district(row)
                row = colorizer(row)
                ow.writerow(row)

        with open(input2) as i2:
            i2r = csv.DictReader(i2)
            for row in i2r:
                row = salesclean(row)
                row = region_district(row)
                row = colorizer(row)
                ow.writerow(row)

        with open(input3) as i3:
            i3r = csv.DictReader(i3)
            for row in i3r:
                row = salesclean(row)
                row = region_district(row)
                row = colorizer(row)
                ow.writerow(row)


    t1 = time.clock()
    return 'Process completed! Duration:', t1-t0

print main()
