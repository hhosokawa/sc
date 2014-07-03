from aux_reader import *
import pandas as pd
import time

""" io """

nnea_df =       pd.read_csv('i/nnea/nnea.csv')
tsr_reps =      csv_dic('i/nnea/tsr reps.csv')

divdistrict =   csv_dic('auxiliary/div-district.csv')
divregion =     csv_dic('auxiliary/div-region.csv')

current_time =  time.strftime("%Y-%m-%d")
output =        'P:/_HHOS/Microsoft/NNEA Summary - %s.xls' % (current_time)

############### utils ###############

def ob_tsr(series):
    rep = str(int(series['Master OB Rep']))
    if rep in tsr_reps: return 'TSR'
    else: return 'OB'

def region(series):
    divloc = (str(int(series['Master Division'])) +
              str(int(series['Master Divloc'])))
    return divregion.get(divloc, '')

def district(series):
    divloc = (str(int(series['Master Division'])) +
              str(int(series['Master Divloc'])))
    return divdistrict.get(divloc, '')

def apply_new_fields():
    nnea_df['Region'] = nnea_df.apply(region, axis = 1)
    nnea_df['OB / TSR'] = nnea_df.apply(ob_tsr, axis = 1)
    nnea_df['District'] = nnea_df.apply(district, axis = 1)

def write_xls():
    writer = pd.ExcelWriter(output)
    nnea_df.to_excel(writer, 'nnea_data', index = False)

    # Pivot Table
    nnea_pt = pd.pivot_table(nnea_df,
                             values='Contract Number',
                             rows=['Region', 'District'],
                             cols=['OB / TSR'],
                             aggfunc=lambda x: len(x.unique()))
    nnea_pt.to_excel(writer, 'nnea_summary')
    writer.save()

############### nnea_main() ###############

if __name__ == '__main__':
    t0 = time.clock()
    apply_new_fields()
    write_xls()
    t1 = time.clock()
    print 'nnea_main() complete. Duration:', t1-t0