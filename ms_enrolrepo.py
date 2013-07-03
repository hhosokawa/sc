import time
import pandas as pd
from aux_reader import *

output    = 'o\\Enrol Repo - 2013-06-13.csv'
repo_df   = pd.read_csv('i\\enrol_repo\\contract repo.csv')
tsr_dict  = csv_dic('i\\enrol_repo\\tsr reps.csv')

divregion = csv_dic('auxiliary\\div-region.csv')

#################################################################################
## Function Definitions

def region(series):
    divloc = (str(int(series['Master Division'])) +
              str(int(series['Master Divloc'])))
    return divregion.get(divloc, '')

def OB_TSR(series):
    rep = str(int(series['Master OB Rep']))
    if rep in tsr_dict: return 'TSR'
    else: return 'OB'

#################################################################################
## Main

def main():
    t0 = time.clock()

    # Fill Missing Values
    repo_df['Master Division'].fillna(0, inplace = True)
    repo_df['Master Divloc'].fillna(0, inplace = True)
    repo_df['Master OB Rep'].fillna(0, inplace = True)
    repo_df['Master-Master Number'].fillna(repo_df['Master Number'],
                                           inplace = True)
    repo_df['Master-Master Name'].fillna(repo_df['Master Name'],
                                           inplace = True)

    # Fill Region and TSR / OB Columns
    repo_df['Region'] = repo_df.apply(region, axis = 1)
    repo_df['OB / TSR'] = repo_df.apply(OB_TSR, axis = 1)

    # Drop Duplicates
    repo_df.drop_duplicates(cols='Contract Number', take_last=True, inplace=True)

    # Write to CSV
    repo_df.to_csv(output, index=False)
    t1 = time.clock()
    return 'Process completed! Duration:', t1-t0

print main()
