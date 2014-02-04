import time
import pandas as pd
from aux_reader import *

""" io """

repo_df   =     pd.read_csv('i\\enrol_repo\\contract_repo.csv')
tsr_dict  =     csv_dic('i\\enrol_repo\\tsr reps.csv')
divregion =     csv_dic('auxiliary\\div-region.csv')
divdistrict =   csv_dic('auxiliary\\div-district.csv')

output    =     'o\\Enrol Repo - %s.csv' % (time.strftime("%Y-%m-%d"))

############### Uitls ###############

def region(series):
    divloc = (str(int(series['Master Division'])) +
              str(int(series['Master Divloc'])))
    return divregion.get(divloc, '')

def district(series):
    divloc = (str(int(series['Master Division'])) +
              str(int(series['Master Divloc'])))
    return divdistrict.get(divloc, '')

def OB_TSR(series):
    rep = str(int(series['Master OB Rep']))
    if rep in tsr_dict: return 'TSR'
    else: return 'OB'

############### enrol_repo_main() ###############

if __name__ == '__main__':
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
    repo_df['District'] = repo_df.apply(district, axis = 1)
    repo_df['OB / TSR'] = repo_df.apply(OB_TSR, axis = 1)

    # Drop Duplicates
    repo_df.drop_duplicates(cols='Contract Number', take_last=True, inplace=True)

    # Write to CSV
    repo_df.to_csv(output, index=False)
    t1 = time.clock()
    print 'enrol_repo_main() completed. Duration:', t1-t0

