import csv
import time
from aux_reader import *

#################################################################################
## Function Definitions
output = 'o\\29-Jan-13 - MS Summary.csv'
input1 = 'i\\ms_summary\\2012-2013 Yoy Raw Data - Referrals - 29-Jan-13.csv'
input2 = 'i\\ms_summary\\2012-2013 Yoy Raw Data - Sales - 29-Jan-13.csv'

FX = csv_dic('auxiliary\\dictFX.csv')
ref_agreementcat = csv_dic('auxiliary\\ref-agreementcategory.csv')

#################################################################################
##  Dictionary Pictionary Jars

majoraccts = csv_dic('auxiliary\\enrol - major customers.csv')
otherdiv = {'200':'100', '100':'200'}
venprogram = csv_dic('auxiliary\\dictvenprograms.csv')
divregion = csv_dic('auxiliary\\div-region.csv')
ref_revtype = csv_dic('auxiliary\\ref-revtype.csv')

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
    row['Custom Category'] = row['Referral Source']
    row['Custom Item Category'] = ''

    # ESA 3.0 Classification
    if row['Referral Source'] == 'ESA 3.0':
        
        # New Contracts VS Renewal Contracts
        if ref_agreementcat[row['Referral Number']] == 'New Contract':
            row['Custom Sub Category'] = 'New Contract'
        else:
            row['Custom Sub Category'] = 'Renew Contract'
        if row['Referral Rev Type'] == 'N': 
            row['Custom Sub Category'] = 'New Contract'
        elif row['Referral Rev Type'] == 'RN': 
            row['Custom Sub Category'] = 'Renew Contract'

        # Major VS Corporate
        if row['Referral Rev Type'] in ref_revtype: 
            row['Custom Item Category'] = ref_revtype[row['Referral Rev Type']]
        else: 
            row['Custom Item Category'] = 'Other'

        if row['Customer Number'] in majoraccts:
            row['Custom Category'] = 'ESA 3.0 - MAJOR'
        else: row['Custom Category'] = 'ESA 3.0 - CORPORATE'

        if (row['Product Item Desc'] == 'CORPORATE ACCOUNT' or
            row['Product Item Desc'] == 'MAJOR ACCOUNT' or
            row['Product Item Bill Type'] in ['ADOT', 'TRUP']): pass
        else: row['Imputed Revenue'] = 0

    # ESA 2.0 Classification
    elif row['Referral Source'] == 'ESA 2.0':
        if row['Product Item Bill Type'] == 'TRUP':
            row['Custom Sub Category'] = 'EA True Up'
        elif row['Product Item Bill Type'] == 'ADOT': 
            row['Custom Sub Category'] = 'EA Add On'
        elif row['Product Item Bill Type'] == 'OAP2': 
            row['Custom Sub Category'] = 'Year 2'
        elif row['Product Item Bill Type'] == 'OAP3': 
            row['Custom Sub Category'] = 'Year 3'
        elif row['Product Item Bill Type'] == 'O':
            if ('Y1' or 'YR1') in row['Product Item Desc']: 
                row['Custom Sub Category'] = 'Year 1'
            elif ('Y2' or 'YR2') in row['Product Item Desc']: 
                row['Custom Sub Category'] = 'Year 2'
            elif ('Y3' or 'YR3') in row['Product Item Desc']: 
                row['Custom Sub Category'] = 'Year 3'
            else: 
                row['Custom Sub Category'] = 'Other'
        else: 
            row['Custom Sub Category'] = 'Year 1'

    # MS SIP Classification
    elif row['Referral Source'] == 'MS SIP':
        row['Custom Sub Category'] = row['Product Item Desc'].title()

    # SPLA Classification
    elif ('SPLA' and 'FENCED DEAL') in row['Referral Notes']:
        row['Custom Sub Category'] = 'SPLA'

    # Other Classification
    else: row['Custom Sub Category'] = 'Other'

    # Ref PO 'CR' = 0 Rule
    if row['Referral PO Reference'] == 'CR':
        row['Imputed Revenue'] = 0
        row['Gross Revenue'] = 0
    return row

# Cleans Sales Data
def salesclean(row):
    row['Revenue Type'] = 'Sales'
    row['Imputed Revenue'] = row['Gross Revenue']
    row['Custom Category'] = 'NON-EA DIRECT'
    row['Custom Sub Category'] = venprogram.get(row['Item Prodtype Venprogram'],
                                                'EA Indirect and Other')
    return row

# Virtual Adj
def virtadj(row):
    if row['Divloc'] == '99':
        d, p, y = row['Division'], row['Period'], row['Year']
        row['GP'] = float((row['GP'])) * adjFX(d, p, y)
        row['Gross Revenue'] = float(row['Gross Revenue']) * adjFX(d, p, y)
        row['Imputed Revenue'] = float(row['Imputed Revenue']) * adjFX(d, p, y)
        row['Division'] = otherdiv[d]
    return row

# Region Adjustment
def region(row):
    divloc = str(row['Division']) + str(row['Divloc'])
    row['Region'] = divregion.get(divloc, 'N/A')
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
                row = region(row)
                ow.writerow(virtadj(row))

        with open(input2) as i2:
            i2r = csv.DictReader(i2)
            ow = csv.DictWriter(o, fieldnames=header)
            for row in i2r:
                row = salesclean(row)
                row = region(row)
                ow.writerow(virtadj(row))

    t1 = time.clock()
    return 'Process completed! Duration:', t1-t0

print main()
