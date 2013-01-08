import csv
import time

# Converts CSV to dict
def csv_dic(filename):
    reader = csv.reader(open(filename, "rb"))
    my_dict = dict((k, v) for k, v in reader)
    return my_dict

#################################################################################
## Function Definitions
output = 'o\\31-Dec-12 - MS Summary.csv'
input1 = 'i\\2011-2012 Yoy Raw Data - Referral - MS 31-Dec-12.csv'
input2 = 'i\\2011-2012 Yoy Raw Data - Indirect - MS 31-Dec-12.csv'

FX = csv_dic('auxiliary\\dictFX.csv')
ref_agreementcat = csv_dic('auxiliary\\ref-agreementcategory.csv')

#################################################################################
##  Dictionary Pictionary Jars

majoraccts = csv_dic('auxiliary\\enrol - major customers.csv')
otherdiv = {'200':'100', '100':'200'}
venprogram = csv_dic('auxiliary\\dictvenprograms.csv')
divregion = csv_dic('auxiliary\\div-region.csv')
ref_revtype = csv_dic('auxiliary\\ref-revtype.csv')

#################################################################################
## Function Definitions

# Adjusts for Finance's FX File
def adjFX(div, period, year):
    if div == 100: i = 'CDN-US-'
    else: i = 'US-CDN-'
    answer = i+str(period)+'-'+str(year)
    return float(FX[answer])

# Clean Referrals Data
def refclean(row):
    row['Quarter'] = 'Q' + str(row['Quarter'])
    row['Revenue Type'] = 'Referrals'
    row['Gross Revenue'] = row['GP']
    row['Custom Category'] = row['Referral Source']
    row['Custom Item Category'] = ''
    csc = 'Custom Sub Category'
    pibt = 'Product Item Bill Type'
    rrt = 'Referral Rev Type'
    cic = 'Custom Item Category'
    pid = 'Product Item Desc'

    # ESA 3.0 Classification
    if row['Referral Source'] == 'ESA 3.0':
        if ref_agreementcat[row['Referral Number']] == 'New Contract':
            row[csc] = 'New Contract'
        else:
            row[csc] = 'Renew Contract'
        if row[rrt] == 'N': row[csc] = 'New Contract'
        elif row[rrt] == 'RN': row[csc] = 'Renew Contract'
        if row[rrt] in ref_revtype: row[cic] = ref_revtype[row[rrt]]
        else: row[cic] = 'Other'
        if row['Customer Number'] in majoraccts:
            row['Custom Category'] = 'ESA 3.0 - MAJOR'
        else: row['Custom Category'] = 'ESA 3.0 - CORPORATE'
        if (row[pid] == 'CORPORATE ACCOUNT' or
            row[pid] == 'MAJOR ACCOUNT' or
            row[pibt] in ['ADOT', 'TRUP']): pass
        else: row['Imputed Revenue'] = 0

    # ESA 2.0 Classification
    elif row['Referral Source'] == 'ESA 2.0':
        if row[pibt] == 'TRUP': row[csc] = 'EA True Up'
        elif row[pibt] == 'ADOT': row[csc] = 'EA Add On'
        elif row[pibt] == 'OAP2': row[csc] = 'Year 2'
        elif row[pibt] == 'OAP3': row[csc] = 'Year 3'
        elif row[pibt] == 'O':
            if ('Y1' or 'YR1') in row[pid]: row[csc] = 'Year 1'
            elif ('Y2' or 'YR2') in row[pid]: row[csc] = 'Year 2'
            elif ('Y3' or 'YR3') in row[pid]: row[csc] = 'Year 3'
            else: row[csc] = 'Other'
        else: row[csc] = 'Year 1'

    # MS SIP Classification
    elif row['Referral Source'] == 'MS SIP':
        row[csc] = row[pid].title()

    # Other Classification
    else: row[csc] = 'Other'

    # Ref PO 'CR' = 0 Rule
    if row['Referral PO Reference'] == 'CR':
        row['Imputed Revenue'] = 0
        row['Gross Revenue'] = 0
    return row

# Cleans Sales Data
def salesclean(row):
    row['Revenue Type'] = 'Sales'
    row['Imputed Revenue'] = row['Gross Revenue']
    row['Custom Category'] = 'NON-EA DIRECT'
    ipv = row['Item Prodtype Venprogram']
    try: row['Custom Sub Category'] = venprogram[ipv]
    except: row['Custom Sub Category'] = 'EA Indirect and Other'
    return row

# Virtual Adj
def virtadj(row):
    if row['Divloc'] == '99':
        d, p, y = row['Division'], row['Period'], row['Year']
        row['GP'] = float((row['GP'])) * adjFX(d, p, y)
        row['Gross Revenue'] = float(row['Gross Revenue']) * adjFX(d, p, y)
        row['Imputed Revenue'] = float(row['Imputed Revenue']) * adjFX(d, p, y)
        row['Division'] = otherdiv[d]
    return row

# Region Adjustment
def region(row):
    divloc = str(row['Division']) + str(row['Divloc'])
    try: row['Region'] = divregion[divloc]
    except: row['Region'] = 'N/A'
    return row


#################################################################################
## Main Program

def main():
    t0 = time.clock()
    header = set()

    # Get all input headers -> output header
    with open(input1) as i1: header.update(csv.DictReader(i1).fieldnames)
    with open(input2) as i2: header.update(csv.DictReader(i2).fieldnames)
    header = tuple(header)

    # Analyze Input 1-2 -> Output
    with open(output, 'wb') as o:
        writer = csv.writer(o)
        writer.writerows([header])
        with open(input1) as i1:
            i1r = csv.DictReader(i1)
            ow = csv.DictWriter(o, fieldnames=header)
            for row in i1r:
                row = refclean(row)
                row = region(row)
                ow.writerow(virtadj(row))

        with open(input2) as i2:
            i2r = csv.DictReader(i2)
            ow = csv.DictWriter(o, fieldnames=header)
            for row in i2r:
                row = salesclean(row)
                row = region(row)
                ow.writerow(virtadj(row))

    t1 = time.clock()
    return 'Process completed! Duration:', t1-t0

print main()
