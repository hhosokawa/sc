import os
import time
import numpy as np
import pandas as pd
from aux_reader import *

input = 'i/mda/mda 2012-2013 q2.csv'
output = 'o/2012-2013 Q2 MDA.csv'

gov_edu = ['FED', 'STAT', 'PROV', 'MUN', 'EDUC', 
           'SSLP', 'QFED', 'GRIC', 'SFGP']

#################################################################################
## Function Definitions

def cust_type(series):

    # SMB / GOVT & EDUC / ENTERPRISE
    pc_count = series['Master PC Count']
    acct_type = series['Bill To Account Type']
    if acct_type in gov_edu:
        return 'GOVT & EDUC'
    else:
        if pc_count < 2000: return 'SMB'
        else: return 'ENTERPRISE'

#################################################################################
## Main

def main():
    t0 = time.clock()
    df = pd.read_csv(input)

    # Remove Duplicates / Customer Type / Write to .csv
    df.drop_duplicates(cols='Order Line ID')
    df['Cust Type'] = df.apply(cust_type, axis = 1)
    df.to_csv(output)

    return 'Process completed! Duration:', time.clock()-t0

print main()

