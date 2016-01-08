import time
import pandas as pd
import csv

# io
ENROL_DF = pd.read_csv('i\\enrol_customer\\enrol.csv',
                       encoding = "ISO-8859-1")
OUTPUT = 'o\\Enrol Cust - %s.xls' % (time.strftime("%Y-%m-%d"))

# Converts Csv -> Dictionary
def csv_dic(filename, col = 1):
    reader = csv.reader(open(filename, "rb"))
    if col == 1:
        my_dict = dict((k, v) for k, v in reader)
    elif col == 2:
        my_dict = dict((k, (v0, v1)) for k, v0, v1 in reader)
    return my_dict

# Pictionary
divregion = csv_dic('i\\enrol_customer\\div-region.csv')
divdistrict = csv_dic('i\\enrol_customer\\div-district.csv')
masterid_region_district = csv_dic('i\\enrol_customer\\masterid_region_district.csv', 2)
tsr_dict = csv_dic('i\\enrol_customer\\tsr reps.csv')

############### utils ###############

def region(series):
    try:
        master_number = str(int(series['Master Number']))
    except ValueError:
        master_number = ''
    if master_number in masterid_region_district:
        return masterid_region_district[master_number][0]
    else:
        divloc = (str(int(series['Master Division'])) +
        str(int(series['Master Divloc'])))
        return divregion.get(divloc, '')

def district(series):
    try:
        master_number = str(int(series['Master Number']))
    except ValueError:
        master_number = ''
    if master_number in masterid_region_district:
        return masterid_region_district[master_number][1]
    else:
        divloc = (str(int(series['Master Division'])) +
        str(int(series['Master Divloc'])))
        return divdistrict.get(divloc, '')

def OB_TSR(series):
    rep = str(int(series['Master OB Rep']))
    if rep in tsr_dict: return 'TSD'
    else: return 'OB'

def lic_prog_name(series):
    if series['Licensing ProgramName'] == 'MSEA': return 'MSEA'
    elif series['Licensing ProgramName'] == 'MSACAD': return 'MSEA'
    else: return 'MSOTHER'

def unique_id(series):
    unique_id = (str(int(series['Master-Master Number'])) +
                 str(series['Licensing ProgramName']))
    return unique_id

############### enrol_cust_main() ###############

if __name__ == '__main__':
    t0 = time.clock()

    # 1) Enrollment DataFrame 

    #    Fill Missing Values
    ENROL_DF['Master OB Rep'].fillna(0, inplace = True)
    ENROL_DF['Master Divloc'].fillna(0, inplace = True)
    ENROL_DF['Master Division'].fillna(0, inplace = True)
    ENROL_DF['Master-Master Name'].fillna(ENROL_DF['Master Name'],
                                         inplace = True)
    ENROL_DF['Master-Master Number'].fillna(ENROL_DF['Master Number'],
                                           inplace = True)
    ENROL_DF['Master-Master Number'].fillna(0, inplace = True)

    #    Fill District, OB/TSR, and Region columns
    ENROL_DF['District'] = ENROL_DF.apply(district, axis = 1)
    ENROL_DF['OB / TSR'] = ENROL_DF.apply(OB_TSR, axis = 1)
    ENROL_DF['Region'] = ENROL_DF.apply(region, axis = 1)

    #    Drop Duplicates
    ENROL_DF.drop_duplicates(cols='Contract Number', take_last=True,
                            inplace=True)

    # 2) Customer DataFrame 
    cust_df = ENROL_DF.copy()
    
    #    Fill Unique ID, MSEA
    cust_df['Unique ID'] = cust_df.apply(unique_id, axis = 1)
    cust_df['Licensing ProgramName'] = cust_df.apply(lic_prog_name, axis = 1)

    #    Drop Duplicates
    cust_df.drop_duplicates(cols='Unique ID', take_last=True, inplace=True)

    # 3) Write to Excel 
    writer = pd.ExcelWriter(OUTPUT)
    cust_df.to_excel(writer, 'cust_data', index = False)
    ENROL_DF.to_excel(writer, 'enrol_data', index = False)
    writer.save()
    t1 = time.clock()
    print 'enrol_cust_main() completed. Duration:', t1-t0