import os
import time
import numpy as np
import pandas as pd
from aux_reader import *

input = 'i/database_revrec'
output = 'C:/Portable Python 2.7/App/Scripts/o/recurring_revenue.csv'

mm_name = csv_dic('C:/Portable Python 2.7/App/Scripts/auxiliary/mastermaster_number_name.csv')
divregion = csv_dic('C:/Portable Python 2.7/App/Scripts/auxiliary/div-region.csv')
recurringrevenue = ['SWMTN', 'SWSUB', 'SWBNDL', 'CLOUD']

#################################################################################
## Function Definitions

def region(series):
    # Determines Region
    divloc = (str(int(series['Softchoice Division'])) +
              str(int(series['Softchoice DivLoc'])))
    return divregion.get(divloc, '')


def ref_clean(df):
    # fillna - Referal Fee Received Total & Master Master Number
    df['Referral Fee Received Total'].fillna(
        df['Referral Net Expected Fee Total'], inplace = True)

    # Rename Columns
    df.rename(columns={
        'Softchoice Master Master Number' : 'Softchoice Bill to Master Master Number',
        'Master NAICS Description': 'Bill to Master NAICS Description',
        'SoftChoice Master Name' : 'Bill to Master Name',
        'SoftChoice Master Number' : 'Bill to Master Number',
        'Vendor Name' : 'Product Publisher Name',
        'Referral Fiscal Year' : 'Invoice date (SC FY)',
        'Referral Fiscal Quarter': 'Invoice date (SC FQ)',
        'Referral Fiscal Period': 'Invoice date (SC FM)',
        'Referral Division': 'Softchoice Division',
        'Referral Divloc Number': 'Softchoice DivLoc',
        'Referral Super Category' : 'Item Super Category',
        'Referral Fee Received Total' : 'Total GP $'}, inplace = True)
    return df


def sales_clean(series):
    # Determine Run-Rate / Project / Recurring Revenue
    if series['Item Revenue Recognition'] in recurringrevenue:
        return 'Recurring Revenue'
    else:
        gp = float(series['Total GP $'])
        if -1000 <= gp <= 1000: return 'Run-Rate'
        else: return 'Project'

def master_id(series):
    # Determines correct Master Master Name
    mm_number = series['Softchoice Bill to Master Master Number']
    master_name = series['Bill to Master Name']
    if mm_number in mm_name:
        return mm_name[mm_number]
    else:
        return master_name

#################################################################################
## Main

def main():
    t0 = time.clock()
    # .csv Merge
    master_df = pd.DataFrame()
    os.chdir(input)
    for files in os.listdir("."):
        df = pd.read_csv(files)

        # Data Scrub Referrals / Sales
        if 'ref' in files: df = ref_clean(df)
        else: df['Sales Type'] = df.apply(sales_clean, axis = 1)

        # Determines correct Master Master Name
        df['Master ID'] = df.apply(master_id, axis = 1)

        # df Select Headers
        filter_df = df.ix[:, [
                    'Softchoice Division',
                    'Softchoice DivLoc',
                    'Invoice date (SC FY)',
                    'Sales Type',
                    'Master ID',
                    'Total GP $']]

        # Filter Region / Groupby / Conatenate to Master
        filter_df['Region'] = filter_df.apply(region, axis = 1)
        filter_df = filter_df.drop(['Softchoice DivLoc',
                                    'Softchoice Division'], axis = 1)
        group_df = filter_df.groupby([
                        'Region',
                        'Sales Type',
                        'Master ID',
                        'Invoice date (SC FY)']).sum()
        master_df = pd.concat([group_df, master_df])

    # Master_df -> output.csv
    master_df.to_csv(output)
    return 'Process completed! Duration:', time.clock()-t0

print main()

"""
    master_df = master_df.reset_index()
    master_df = pd.pivot_table(master_df,
                values = 'Total GP $',
                rows = ['Region', 'Sales Type'],
                cols = ['Invoice date (SC FY)'],
                aggfunc=np.sum)
"""
