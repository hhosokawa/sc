import csv
import time
from aux_reader import *

############### io ###############

output = 'o\\ms summary.csv'
input1 = 'i\\ms_summary\\ref.csv'
input2 = 'i\\ms_summary\\sales.csv'
input3 = 'i\\ms_summary\\bi.csv'

############### pictionary ###############

order = {}
majoraccts = csv_dic('auxiliary\\enrol - major customers.csv')
venprogram = csv_dic('auxiliary\\dictvenprograms.csv')
ref_revtype = {'AO': 'EA Add-On',
               'N': 'EA New',
               'RC': 'EA Recurring',
               'RN': 'EA Renewal',
               'TU': 'EA True-Up'}

############### utils ###############

# Obtain Headers
def get_header():
    header = set()
    with open(input3) as i3: header.update(csv.DictReader(i3).fieldnames)
    new_fields = set(['Category A', 'Category B', 'Category C'])
    header = new_fields | header
    try: header.remove('')
    except KeyError: pass
    header = tuple(sorted(header, key=lambda item: item[0]))
    return header

# Referral Clean
def scan_referrals():
    with open(input1) as i1:
        for r in csv.DictReader(i1): refclean(r)

def refclean(r):
    # Category B Classification
    catB = r['Referral Source']
    catC = ''

    # ESA 3.0 Classification
    if r['Referral Source'] == 'ESA 3.0':

        # Major VS Corporate
        catA = 'EA'
        catC = ref_revtype.get(r['Referral Rev Type'], 'Other')

        if (r['SoftChoice Customer Number'] in majoraccts
        or 'MAJOR' in r['Referral Notes']
        or r['Product Item Desc'] == 'HELP DESK INCENTIVE'):
            catB = 'ESA 3.0 - MAJOR'
        else:
            catB = 'ESA 3.0 - CORPORATE'

    # ESA 2.0 Classification
    elif r['Referral Source'] == 'ESA 2.0':
        catA = 'EA'
        if r['Product Item Bill Type'] == 'TRUP':   catC = 'EA True Up'
        elif r['Product Item Bill Type'] == 'ADOT': catC = 'EA Add On'
        elif r['Product Item Bill Type'] == 'OAP2': catC = 'Year 2'
        elif r['Product Item Bill Type'] == 'OAP3': catC = 'Year 3'
        elif r['Product Item Bill Type'] == 'O':
            if ('Y1' or 'YR1') in r['Product Item Desc']:   catC = 'Year 1'
            elif ('Y2' or 'YR2') in r['Product Item Desc']: catC = 'Year 2'
            elif ('Y3' or 'YR3') in r['Product Item Desc']: catC = 'Year 3'
            else: catC = 'Other'
        else: catC = 'Year 1'

    # MS SIP Classification
    elif r['Referral Source'] == 'MS SIP':
        catA = 'EA'
        catC = r['Product Item Desc'].title()

    # SPLA Classification
    elif ('SPLA' and 'FENCED DEAL') in r['Referral Notes']:
        catA = 'SPLA'
        catB = 'Referrals'

    # Other Classification
    else: catA = 'Other'

    # Absorb into Dictionary
    order[r['Referral Number']] = (catA, catB, catC)
    return


# Sales Clean
def scan_sales():
    with open(input2) as i2:
        for r in csv.DictReader(i2): 
            salesclean(r)

def salesclean(r):
    ven = venprogram.get(r['Item Prodtype Venprogram'],
                         'EA Indirect and Other')
    if ven in ['Open', 'Select']:
        catA = ven
        catB, catC = '', ''
    elif ven == 'SPLA':
        catA = ven
        catB, catC = 'Sales', ''
    elif ven == 'EA Indirect and Other':
        catA = 'EA'
        catB, catC = ven, ''

    # Absorb into Dictionary
    order[r['Order Number']] = (catA, catB, catC)
    return

# Insert Category for Order
def add_cat(r):
    if r['Order Number'] in order:
        r['Category A'] = order[r['Order Number']][0]
        r['Category B'] = order[r['Order Number']][1]
        r['Category C'] = order[r['Order Number']][2]
    else:
        r['Category A'] = 'N/A'
        r['Category B'] = 'N/A'
        r['Category C'] = 'N/A'
    return r

############### main ###############

def ms_summary_main():
    t0 = time.clock()
    header = get_header()

    # Scan Sales / Referral Data -> Dictionary
    scan_referrals()
    scan_sales()

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
    print 'ms_summary_main() completed! Duration:', t1-t0
    return

if __name__ == '__main__':
    ms_summary_main()


