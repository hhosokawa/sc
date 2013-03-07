import csv
import time
from aux_reader import *

################ Inputs ################
output = 'o\\bi_class.csv'
refinput = 'i\\bi_class\\ref-custom cat.csv'
salesinput = 'i\\bi_class\\sales-custom cat.csv'

ref_tree = tree()
sales_tree = tree()
ref_revtype = {'AO': 'EA Add-On', 'N': 'EA New', 'RC': 'EA Recurring',
               'RN': 'EA Renewal', 'TU': 'EA True-Up'}
ref_revtype2 = {'TRUP': 'EA True-Up', 'ADOT': 'EA Add-On', 
                'OAP2': 'Year 2', 'OAP3': 'Year 3'}
venprogram = csv_dic('auxiliary\\dictvenprograms.csv')

################ Funcs ################

def refclean(r):
    x = ref_tree[r['Referral Number']] 
    x['Referral #'] = r['Referral Number']

    # ESA 3.0
    x['Category A'] = r['Referral Source']
    if x['Category A'] == 'ESA 3.0':
        if r['MS Major Acct (Y/N)'] == 'Y':
            x['Category A'] = 'ESA 3.0 - MAJOR'        
        else:
            x['Category A'] = 'ESA 3.0 - CORPORATE'
        x['Category B'] = ref_revtype.get(r['Referral Rev Type'], 'Other')

    # ESA 2.0
    elif x['Category A'] == 'ESA 2.0':

        if r['Product Item Bill Type'] in ref_revtype2:
            x['Category B'] = ref_revtype2[r['Product Item Bill Type']]
        else:
            if r['Product Item Bill Type'] == 'O':
                if ('Y1' or 'YR1') in r['Product Item Desc']: 
                    x['Category B'] = 'Year 1'
                elif ('Y2' or 'YR2') in r['Product Item Desc']: 
                    x['Category B'] = 'Year 2'
                elif ('Y3' or 'YR3') in r['Product Item Desc']: 
                    x['Category B'] = 'Year 3'
                else: 
                    x['Category B'] = 'Other'
            else:
                x['Category B'] = 'Other'

    # MS SIP
    elif x['Category A'] == 'MS SIP':
        x['Category B'] = r['Product Item Desc'].title()

    # SPLA
    elif ('SPLA' and 'FENCED DEAL') in r['Referral Notes']:
        x['Category B'] = 'SPLA'        

    # OTHER
    elif x['Category A'] == 'OTHER':
        x['Category B'] = 'Other'

    # BLANK REFERRAL SOURCE
    else: 
        x['Category A'] == 'OTHER'
        x['Category B'] = 'Other'
    return

def salesclean(r):
    x = sales_tree[r['Order Number']] 
    x['SO #'] = r['Order Number']
    x['Category A'] = 'NON-EA DIRECT'
    x['Category B'] = venprogram.get(r['Item Prodtype Venprogram'],
                                     'EA Indirect and Other')
    return

################ Main ################

def main():
    t0 = time.clock()    

    # Referral #: Custom Category A-C
    with open(refinput) as i1:
        for r in csv.DictReader(i1): refclean(r)
                
    # Sales #: Custom Category A-C
    with open(salesinput) as i2:
        for r in csv.DictReader(i2): salesclean(r)

    # Output Writer
    header = ('Referral #', 'SO #', 'Category A', 'Category B')
    with open(output, 'wb') as o:
        writer = csv.writer(o)
        ow = csv.DictWriter(o, fieldnames=header)
        writer.writerows([header])
        for ref in ref_tree: ow.writerow(ref_tree[ref])
        for so in sales_tree: ow.writerow(sales_tree[so])
        
    t1 = time.clock()
    print 'Process completed! Duration:', t1-t0
    return

if __name__ == '__main__':
    print main()

