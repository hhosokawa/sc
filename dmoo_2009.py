import csv
import time
import numpy as np
import pandas as pd

input1 = 'i/dmoo_2009/sales_ref - 2009.csv'
output = 'o/2009 net rev.csv'

netrevrecs = ['SWSUB', 'SWMTN', 'HWMTN', 'TRAIN', 'CLOUD']

#################################################################################
## Function Definitions

def rev_rec(series):
    pub = series['Product Publisher Name']
    rev = series['Rev Rec']
    div = int(series['Div'])
    gp = float(series['GP'])
    gross = float(series['Gross Rev'])
    cogs = float(series['Gross Rev']) - float(series['GP'])

    # Net Revenue
    if rev in netrevrecs:
        return gp

    # SWBNDL % Net Revenue
    elif rev == 'SWBNDL':
        if div == 100:
            return gross - (cogs * 0.2486)
        elif div == 200:
            return gross - (cogs * 0.2643)
    # Gross Revenue
    else:
        return gross

#################################################################################
## Main

def main():
    t0 = time.clock()
    df = pd.read_csv(input1)

    # Assign Customer Class
    df['Net Rev'] = df.apply(rev_rec, axis = 1)

    # Write to .csv
    df.to_csv(output)
    return 'Process completed! Duration:', time.clock()-t0

print main()

