import time
import pandas as pd
from aux_reader import *

""" io """

output =        'o/nnea.csv'
nnea_df =       pd.read_csv('i/nnea/nnea.csv')

tsr_reps =      csv_dic('i/nnea/tsr reps.csv')
divdistrict =   csv_dic('auxiliary\\div-district.csv')

############### utils ###############

def ob_tsr(series):
    rep = str(int(series['Master OB Rep']))
    if rep in tsr_reps: return 'TSR'
    else: return 'OB'

def district(series):
    divloc = (str(int(series['Master Division'])) +
              str(int(series['Master Divloc'])))
    return divdistrict.get(divloc, '')

def apply_new_fields():
    nnea_df['District'] = nnea_df.apply(district, axis = 1)
    nnea_df['OB / TSR'] = nnea_df.apply(ob_tsr, axis = 1)

############### nnea_main() ###############

if __name__ == '__main__':
    t0 = time.clock()
    apply_new_fields()
    nnea_df.to_csv(output, index=False)
    t1 = time.clock()
    print 'Process completed! Duration:', t1-t0