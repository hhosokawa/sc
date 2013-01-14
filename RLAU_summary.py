import csv
import time
from decimal import Decimal

output = 'o\\RLau - 2012 Summary Dec Only.csv'
input1 = 'i\\RLAU - 2013-Dec Only Referral.csv'
input2 = 'i\\RLAU - 2013-Dec Only Indirect.csv'

#################################################################################
## Function Definitions

# Converts CSV to dict
def csv_dic(filename):
    reader = csv.reader(open(filename, "rb"))
    my_dict = dict((k, v) for k, v in reader)
    return my_dict

# Converts CSV to dict 2
def csv_dic2(filename):
    reader = csv.reader(open(filename, "rb"))
    my_dict = dict((k, (v1, v2, v3)) for k, v1, v2, v3 in reader)
    return my_dict

# Adds New Headers
def header_add(header):
    newfields = ['Net Revenue','COGS', 'Cust Type', '']
    for newfield in newfields: header.add(newfield)
    return header

# Clean Referrals Data
def refclean(row):
    row['Quarter'] = 'Q' + str(row['Quarter'])
    row['Revenue Type'] = 'Referrals'
    row['Gross Revenue'] = row['GP']
    row['Custom Category'] = row['Referral Source']
    row['Custom Item Category'] = ''
    pibt = row['Product Item Bill Type']
    rrt = row['Referral Rev Type']
    pid = row['Product Item Desc']

    # ESA 3.0 Classification
    if row['Referral Source'] == 'ESA 3.0':
        try:
            if ref_agreementcat[row['Referral Number']] == 'New Contract':
                row['Custom Sub Category'] = 'New Contract'
            else:
                row['Custom Sub Category'] = 'Renew Contract'
        except:
            row['Custom Sub Category'] = 'Renew Contract'
        if rrt == 'N': row['Custom Sub Category'] = 'New Contract'
        elif rrt == 'RN': row['Custom Sub Category'] = 'Renew Contract'
        if rrt in ref_revtype: row['Custom Item Category'] = ref_revtype[rrt]
        else: row['Custom Item Category'] = 'Other'
        if row['Customer Number'] in majoraccts:
            row['Custom Category'] = 'ESA 3.0 - MAJOR'
        else: row['Custom Category'] = 'ESA 3.0 - CORPORATE'
        if (pid == 'CORPORATE ACCOUNT' or
            pid == 'MAJOR ACCOUNT' or
            pibt in ['ADOT', 'TRUP']): pass
        else: row['Imputed Revenue'] = 0
    # ESA 2.0 Classification
    elif row['Referral Source'] == 'ESA 2.0':
        if pibt == 'TRUP': row['Custom Sub Category'] = 'EA True Up'
        elif pibt == 'ADOT': row['Custom Sub Category'] = 'EA Add On'
        elif pibt == 'OAP2': row['Custom Sub Category'] = 'Year 2'
        elif pibt == 'OAP3': row['Custom Sub Category'] = 'Year 3'
        elif pibt == 'O':
            if ('Y1' or 'YR1') in pid: row['Custom Sub Category'] = 'Year 1'
            elif ('Y2' or 'YR2') in pid: row['Custom Sub Category'] = 'Year 2'
            elif ('Y3' or 'YR3') in pid: row['Custom Sub Category'] = 'Year 3'
            else: row['Custom Sub Category'] = 'Other'
        else: row['Custom Sub Category'] = 'Year 1'
    # MS SIP Classification
    elif row['Referral Source'] == 'MS SIP':
        row['Custom Sub Category'] = pid.title()
    # Other Classification
    else: row['Custom Sub Category'] = 'Other'

    # Ref PO 'CR' = 0 Rule
    if row['Referral PO Reference'] == 'CR':
        row['Imputed Revenue'] = 0
        row['Gross Revenue'] = 0
        row['Net Revenue'] = 0
    return row

# Cleans Sales Data
def salesclean(row):
    row['Revenue Type'] = 'Sales'
    row['Imputed Revenue'] = row['Gross Revenue']
    row['Custom Category'] = 'SALES'
    ipv = row['Item Prodtype Venprogram']
    try: row['Custom Sub Category'] = venprogram[ipv]
    except: row['Custom Sub Category'] = 'EA Indirect and Other'
    return row

# Region Adjustment
def region(row):
    divloc = str(row['Division']) + str(row['Divloc'])
    try: row['Region'] = divregion[divloc]
    except: row['Region'] = 'N/A'
    return row

def custtype(row):
    try:
        if row['Revenue Type'] == 'Referrals':
            pc_count = int(ref_cus_type[row['Customer Number']][1])
            cus_type = ref_cus_type[row['Customer Number']][2]
        elif row['Revenue Type'] == 'Sales':
            pc_count = row["Softchoice Master # PC's"]
            cus_type = row['Softchoice Master Cust Type']
        else:
            raise RuntimeError('Invalid Revenue Type')
    except:
        pc_count = 0
        cus_type = 'N/A'
    if cus_type in (['FED', 'STAT', 'PROV', 'MUN', 'EDUC', 'SSLP',
                     'QFED', 'GRIC', 'SFGP']):
        row['Cust Type'] = 'GOVT & EDUC'
    else:
        if pc_count < 2000: row['Cust Type'] = 'SMB'
        else: row['Cust Type'] = 'Enterprise'
    return row

def netrev(row):
    grossrev = Decimal(row['Gross Revenue'])
    gp = Decimal(row['GP'])
    if row['Revenue Type'] == 'Referrals':
        row['Net Revenue'] = grossrev
    elif row['Revenue Type'] == 'Sales':
        # GROSS
        if row['Item Revenue Recognition'] in ['BOOK', 'DSRV', 'HW',
        'IHADJ', 'ISADJ','IVADJ','PSRV', 'SW', 'SWLIC']:
            row['Net Revenue'] = grossrev
        # NET
        elif row['Item Revenue Recognition'] in ['HWMTN', 'SWMTN',
        'SWSUB', 'TRAIN']:
            row['Net Revenue'] = gp
        # SWBUNDL
        elif row['Item Revenue Recognition'] == 'SWBNDL':
            if row['Division'] == '100':
                if row['Product Publisher Name'] == 'IBM': pct = Decimal(0.2)
                elif row['Product Publisher Name'] == 'MCAFEE INC': pct = Decimal(0.1854)
                elif row['Product Publisher Name'] == 'MICROSOFT': pct = Decimal(0.2754)
                elif row['Product Publisher Name'] == 'QUEST SOFTWARE': pct = Decimal(0.1968)
                elif row['Product Publisher Name'] == 'SOPHOS': pct = Decimal(0.1221)
                elif row['Product Publisher Name'] == 'SYMANTEC': pct = Decimal(0.2172)
                else: pct = Decimal(0.2486)
            elif row['Division'] == '200':
                if row['Product Publisher Name'] == 'IBM': pct = Decimal(0.2)
                elif row['Product Publisher Name'] == 'MCAFEE INC': pct = Decimal(0.2099)
                elif row['Product Publisher Name'] == 'MICROSOFT': pct = Decimal(0.2856)
                elif row['Product Publisher Name'] == 'QUEST SOFTWARE': pct = Decimal(0.2378)
                elif row['Product Publisher Name'] == 'SOPHOS': pct = Decimal(0.1602)
                elif row['Product Publisher Name'] == 'SYMANTEC': pct = Decimal(0.238)
                else: pct = Decimal(0.2643)
            cogs = grossrev - gp
            row['Net Revenue'] = grossrev - (pct*cogs)
    row['COGS'] = grossrev - gp
    return row

#################################################################################
##  Dictionary Pictionary Jars
majoraccts = csv_dic('auxiliary\\enrol - major customers.csv')
ref_agreementcat = csv_dic('auxiliary\\ref-agreementcategory.csv')

venprogram = csv_dic('auxiliary\\dictvenprograms.csv')
divregion = csv_dic('auxiliary\\div-region.csv')
ref_revtype = csv_dic('auxiliary\\ref-revtype.csv')
so_cus_type = csv_dic2('auxiliary\\Cust SO - PC Count - Type.csv')
ref_cus_type = csv_dic2('auxiliary\\Cust REF - PC Count - Type.csv')

#################################################################################
## Main Program

def main():
    t0 = time.clock()
    header = set()
    # Get all input headers -> output header
    with open(input1) as i1: header.update(csv.DictReader(i1).fieldnames)
    with open(input2) as i2: header.update(csv.DictReader(i2).fieldnames)
    header = header_add(header)
    header = tuple(header)

    # Analyze Input X -> Output
    with open(output, 'wb') as o:
        writer = csv.writer(o)
        writer.writerows([header])
        with open(input1) as i1:
            i1r = csv.DictReader(i1)
            ow = csv.DictWriter(o, fieldnames=header)
            for row in i1r:
                row = refclean(row)
                row = region(row)
                row = custtype(row)
                row = netrev(row)
                ow.writerow(row)

        with open(input2) as i2:
            i2r = csv.DictReader(i2)
            ow = csv.DictWriter(o, fieldnames=header)
            for row in i2r:
                row = salesclean(row)
                row = region(row)
                row = custtype(row)
                row = netrev(row)
                ow.writerow(row)

    t1 = time.clock()
    return 'Process completed! Duration:', t1-t0

print main()

