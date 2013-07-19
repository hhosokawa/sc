import csv
import time
import numpy as np
import pandas as pd

input1 = 'i/mda/mda 2012-2013 q2.csv'
input2 = 'i/mda/master_id-type.csv'
output = 'o/2012-2013 Q2 MDA.csv'

gov_edu = ['FED', 'STAT', 'PROV', 'MUN', 'EDUC', 
           'SSLP', 'QFED', 'GRIC', 'SFGP']
masterid = {}

#################################################################################
## Function Definitions

def master_id_scrape(r):
    id = int(r['Customer Master Number'])
    acct_type = r['Customer Type']
    pc_count = r["Customer # PC's"]
    if pc_count: 
        pc_count = int(pc_count)
    else: 
        pc_count = 0
  
    if acct_type:
        if acct_type in gov_edu:
            masterid[id] = 'GOVT & EDUC'
        elif pc_count > 2000:
            masterid[id] = 'ENTERPRISE'
        else:
            masterid[id] = 'SMB'

def customer_class(series):
    # Assigns Customer Class
    try: id = int(series['Master ID'])
    except ValueError: id = 0
    return masterid.get(id, 'SMB')
    
#################################################################################
## Main

def main():
    t0 = time.clock()
    df = pd.read_csv(input1)

    # Scrape Master_ID
    with open(input2) as i2:
        for r in csv.DictReader(i2): master_id_scrape(r)

    # Assign Customer Class
    df['Cust Class'] = df.apply(customer_class, axis = 1)    
    
    # Write to .csv
    df.to_csv(output)
    return 'Process completed! Duration:', time.clock()-t0

print main()

