import csv
import time
from aux_reader import *

#################################################################################
## Function Definitions
output = 'o\\MS Summary 2007-2013.csv'
input1 = 'i\\ms_summary\\ref - 2007-2013.csv'
input2 = 'i\\ms_summary\\sales - 2007-2013.csv'

FX = csv_dic('auxiliary\\dictFX.csv')
majoraccts = csv_dic('auxiliary\\enrol - major customers.csv')
ref_agreementcat = csv_dic('auxiliary\\ref-agreementcategory.csv', '3b')

#################################################################################
## Dictionary Pictionary Jars

imputedclear = set()
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
        row['Custom Category B'] = ref_agreementcat.get(
                                   row['Referral Number'], 'Renew Contract')
        if row['Custom Category B'] == 'New Contract': pass
        else: row['Custom Category B'] = 'Renew Contract'
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

    # SPLA Classification
    elif ('SPLA' and 'FENCED DEAL') in row['Referral Notes']:
        row['Custom Category A'] = 'SPLA'
        row['Custom Category A'] = 'Referrals'        

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
    if row['Custom Category B'] == 'SPLA':
        row['Custom Category B'] = 'Sales'
        row['Custom Category A'] = 'SPLA'        
    return row

# Virtual Adj
def virtadj(row):
    if row['Divloc'] == '99':
        d, p, y = row['Division'], row['Period'], row['Year']
        if row['GP']: row['GP'] = float((row['GP'])) * adjFX(d, p, y)
        else: row['GP'] = 0
        row['Gross Revenue'] = float(row['Gross Revenue']) * adjFX(d, p, y)
        row['Imputed Revenue'] = float(row['Imputed Revenue']) * adjFX(d, p, y)
        row['Division'] = otherdiv[d]
    return row

# Region / District Adjustment
def region_district(row):
    divloc = str(row['Division']) + str(row['Divloc'])
    row['Region'] = divregion.get(divloc, 'N/A')
    row['District'] = divdistrict.get(divloc, 'N/A')
    return row


#################################################################################
## Main Program

def main():
    t0 = time.clock()
    header = set()

    # Get all input headers -> output header
    with open(input1) as i1: header.update(csv.DictReader(i1).fieldnames)
    with open(input2) as i2: header.update(csv.DictReader(i2).fieldnames)
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
                row = virtadj(row)
                row = region_district(row)
                ow.writerow(row)

        with open(input2) as i2:
            i2r = csv.DictReader(i2)
            ow = csv.DictWriter(o, fieldnames=header)
            for row in i2r:
                row = salesclean(row)
                row = virtadj(row)
                row = region_district(row)
                ow.writerow(row)

    t1 = time.clock()
    return 'Process completed! Duration:', t1-t0

print main()
