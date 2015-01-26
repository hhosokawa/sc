import time
import pandas as pd
from aux_reader import *

""" io """

enrol_df  =     pd.read_csv('i\\enrol_customer\\enrol.csv',
                            encoding = "ISO-8859-1")
tsr_dict  =     csv_dic('i\\enrol_customer\\tsr reps.csv')
divregion =     csv_dic('auxiliary\\div-region.csv')
divdistrict =   csv_dic('auxiliary\\div-district.csv')

output    =     'o\\Enrol Cust - %s.xls' % (time.strftime("%Y-%m-%d"))

############### utils ###############

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

def lic_prog_name(series):
    if series['Licensing ProgramName'] == 'MSEA': return 'MSEA'
    else: return 'MSOTHER'

def unique_id(series):
    unique_id = (str(int(series['Master-Master Number'])) +
                 str(series['Licensing ProgramName']))
    return unique_id

############### enrol_cust_main() ###############

if __name__ == '__main__':
    t0 = time.clock()

    # _______________ Enrollment DataFrame _______________

    # Fill Missing Values
    enrol_df['Master OB Rep'].fillna(0, inplace = True)
    enrol_df['Master Divloc'].fillna(0, inplace = True)
    enrol_df['Master Division'].fillna(0, inplace = True)
    enrol_df['Master-Master Name'].fillna(enrol_df['Master Name'],
                                         inplace = True)
    enrol_df['Master-Master Number'].fillna(enrol_df['Master Number'],
                                           inplace = True)
    enrol_df['Master-Master Number'].fillna(0, inplace = True)

    # Fill Region and TSR / OB Columns
    enrol_df['Region'] = enrol_df.apply(region, axis = 1)
    enrol_df['OB / TSR'] = enrol_df.apply(OB_TSR, axis = 1)
    enrol_df['District'] = enrol_df.apply(district, axis = 1)

    # Drop Duplicates
    enrol_df.drop_duplicates(cols='Contract Number', take_last=True,
                            inplace=True)

    # _______________ Customer DataFrame _______________

    cust_df = enrol_df.copy()

    # Fill Unique ID, MSEA
    cust_df['Unique ID'] = cust_df.apply(unique_id, axis = 1)
    cust_df['Licensing ProgramName'] = cust_df.apply(lic_prog_name, axis = 1)

    # Drop Duplicates
    cust_df.drop_duplicates(cols='Unique ID', take_last=True, inplace=True)

    # _______________ Write to Excel _______________

    writer = pd.ExcelWriter(output)
    cust_df.to_excel(writer, 'cust_data', index = False)
    enrol_df.to_excel(writer, 'enrol_data', index = False)
    writer.save()
    t1 = time.clock()
    print 'enrol_cust_main() completed. Duration:', t1-t0

