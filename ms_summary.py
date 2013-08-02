import csv
import time
from aux_reader import *

#################################################################################
## io

output = 'o\\ms summary - 2012-2013 Q3.csv'
input1 = 'i\\ms_summary\\ref.csv'
input2 = 'i\\ms_summary\\sales.csv'
input3 = 'i\\ms_summary\\bi.csv'

#################################################################################
## Dictionary Pictionary Jars

order = {}
majoraccts = csv_dic('auxiliary\\enrol - major customers.csv')
venprogram = csv_dic('auxiliary\\dictvenprograms.csv')
ref_revtype = {'AO': 'EA Add-On',
               'N': 'EA New',
               'RC': 'EA Recurring',
               'RN': 'EA Renewal',
               'TU': 'EA True-Up'}

#################################################################################
## Function Definitions

# Clean Referrals Data
def refclean(row):
    catA = row['Referral Source']

    # ESA 3.0 Classification
    if row['Referral Source'] == 'ESA 3.0':

        # Major VS Corporate
        catB = ref_revtype.get(row['Referral Rev Type'], 'Other')

        if (row['SoftChoice Customer Number'] in majoraccts
        or 'MAJOR' in row['Referral Notes']
        or row['Product Item Desc'] == 'HELP DESK INCENTIVE'):
            catA = 'ESA 3.0 - MAJOR'
        else:
            catA = 'ESA 3.0 - CORPORATE'

    # ESA 2.0 Classification
    elif row['Referral Source'] == 'ESA 2.0':
        if row['Product Item Bill Type'] == 'TRUP':   catB = 'EA True Up'
        elif row['Product Item Bill Type'] == 'ADOT': catB = 'EA Add On'
        elif row['Product Item Bill Type'] == 'OAP2': catB = 'Year 2'
        elif row['Product Item Bill Type'] == 'OAP3': catB = 'Year 3'
        elif row['Product Item Bill Type'] == 'O':
            if ('Y1' or 'YR1') in row['Product Item Desc']:   catB = 'Year 1'
            elif ('Y2' or 'YR2') in row['Product Item Desc']: catB = 'Year 2'
            elif ('Y3' or 'YR3') in row['Product Item Desc']: catB = 'Year 3'
            else: catB = 'Other'
        else: catB = 'Year 1'

    # MS SIP Classification
    elif row['Referral Source'] == 'MS SIP':
        catB = row['Product Item Desc'].title()

    # SPLA Classification
    elif ('SPLA' and 'FENCED DEAL') in row['Referral Notes']:
        catA = 'SPLA'
        catB = 'Referrals'

    # Other Classification
    else: catB = 'Other'

    # Absorb into Dictionary
    order[row['Referral Number']] = (catA, catB)
    return

# Cleans Sales Data
def salesclean(row):
    catA = 'NON-EA DIRECT'
    catB = venprogram.get(row['Item Prodtype Venprogram'],
                          'EA Indirect and Other')
    if catB == 'SPLA':
        catB = 'Sales'
        catA = 'SPLA'

    # Absorb into Dictionary
    order[row['Order Number']] = (catA, catB)
    return

# Insert Category for Order
def add_cat(r):
    if r['Order Number'] in order:
        r['Category A'] = order[r['Order Number']][0]
        r['Category B'] = order[r['Order Number']][1]
    else:
        r['Category A'] = 'N/A'
        r['Category B'] = 'N/A'
    return r


#################################################################################
## Main Program
def main():
    t0 = time.clock()

    # Header
    header = set()
    with open(input3) as i3: header.update(csv.DictReader(i3).fieldnames)
    new_fields = set(['Category A', 'Category B'])
    header = new_fields | header
    header = tuple(sorted(header, key=lambda item: item[0]))

    # Scan Sales / Referral Data -> Dictionary
    with open(input1) as i1:
        for r in csv.DictReader(i1): refclean(r)
    with open(input2) as i2:
        for r in csv.DictReader(i2): salesclean(r)

    # Output Writer
    with open(output, 'wb') as o:
        writer = csv.writer(o)
        ow = csv.DictWriter(o, fieldnames=header)
        writer.writerows([header])

        # BI Data + Product Catagory Data
        with open(input3) as i3:
            for r in csv.DictReader(i3):
                add_cat(r)
                ow.writerow(r)

    t1 = time.clock()
    return 'Process completed! Duration:', t1-t0

print main()
