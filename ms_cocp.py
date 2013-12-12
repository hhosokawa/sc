import csv
import time
from aux_reader import *
import dateutil.parser as dparser
from collections import defaultdict
from datetime import datetime, timedelta, date

""" io """

output = 'o/NNEA COCP.csv'
COCP_path = 'i/ms_cocp/COCP.csv'
NNEA_path = 'i/ms_cocp/NNEA.csv'
emp_stb = csv_dic('i/ms_cocp/employee-super_title_branch.csv', 3)
enrolprogramdivloc = csv_dic('i/ms_cocp/enrol-program-divloc.csv', 4)
enrolprogram = {}
enroldistrict = {}

divregion = csv_dic('auxiliary\\div-region.csv')
divdistrict = csv_dic('auxiliary\\div-district.csv')
titleOBTSR = csv_dic('auxiliary\\title-OB_TSR.csv')
enrolset, accttypeset, masterset = set(), set(), set()
month_num = {'Jan': '01', 'Feb': '02', 'Mar': '03',
             'Apr': '04', 'May': '05', 'Jun': '06',
             'Jul': '07', 'Aug': '08', 'Sep': '09',
             'Oct': '10', 'Nov': '11', 'Dec': '12'}

############### Uitls ###############

# Headers
def get_header():
    newfields = set(['NNEA', 'Renewal', 'COCP Win', 'COCP Loss', 'District',
                    'Rep Type', 'Acct Type', 'Notif Month'])
    header = set()
    with open(COCP_path) as i1: header.update(csv.DictReader(i1).fieldnames)
    with open(NNEA_path) as i2: header.update(csv.DictReader(i2).fieldnames)
    header = header | newfields
    header = tuple(sorted(header, key=lambda item:item[0]))
    return header

# Split into 2 Dict: Program / District
def split_enrolprogramdivloc():
    for enrol in enrolprogramdivloc:
        prog, div, divloc, _ = enrolprogramdivloc[enrol]
        enrolprogram[enrol] = prog
        if divloc:
            district = divdistrict.get(div+divloc, '')
            enroldistrict[enrol] = district 

# COCP Clean
def scan_COCP(ow):
    with open(COCP_path) as i1:
        i1r = csv.DictReader(i1)

        # COCP Enrollment Cleaning
        for r in i1r:
            enrolset.add(r['Enrollment #'])
            r = clean(r, 'COCPEnrol')
            if 'Region' in r['Region']:
                r['Region'] = r['Region'].replace('Region', '')
                r['Region'] = r['Region'].strip()
            ow.writerow(r)

            # COCP Acct Cleaning
            r['COCP Loss'], r['COCP Win'] = '', ''
            acctid = r['Master #'] + r['Type']
            if acctid not in accttypeset:
                accttypeset.add(acctid)
                r = clean(r, 'COCPAcct')
                ow.writerow(r)

# NNEA Clean
def scan_NNEA(ow):
    with open(NNEA_path) as i2:
        i2r = csv.DictReader(i2)

        # NNEA Enrollment Cleaning
        for r in i2r:
            r['Rep'] = r['Master OB Rep Name']
            r['Master #'] = r['Master Number']
            r['Customer'] = r['Master Name']
            masterdivloc = r['Master Division'] + r['Master Divloc']
            r['Region'] = divregion[masterdivloc]
            r['District'] = divregion.get(masterdivloc, '')
            if r['Region'] == 'Canada': r['Region'] = 'Canada'
            if masterdivloc == '10099': r['Region'] = 'US Corporate'
            elif masterdivloc == '20099': r['Region'] = 'Canada'
            if 'Region' in r['Region']:
                r['Region'] = r['Region'].replace('Region', '')
                r['Region'] = r['Region'].strip()

            if r['Contract Number'] not in enrolset:
                enrolset.add(r['Contract Number'])
                r = clean(r, 'NNEAEnrol')
                ow.writerow(r)

            # NNEA Acct Cleaning
            r['NNEA'], r['Renewal'] = '', ''
            if r['Master Number'] not in masterset:
                masterset.add(r['Master Number'])
                r = clean(r, 'NNEAAcct')
                ow.writerow(r)

def clean(r, datatype):
    if datatype == 'COCPEnrol':
        r['Effective Month'] = r['Effective Month'].strip()
        r['Acct Type'] = 'Enrollment'
        r['Contract Program Name'] = enrolprogram.get(r['Enrollment #'], '')
        r['Notif Month'] = '-'.join([r['Effective Year'], month_num[r['Effective Month']]])
        r['District'] = enroldistrict.get(r['Enrollment #'], '')
        if r['Type'] == 'Win': r['COCP Win'] = r['Enrollment #']
        else: r['COCP Loss'] = r['Enrollment #']
    elif datatype == 'COCPAcct':
        r['Effective Month'] = r['Effective Month'].strip()
        r['Contract Program Name'] = ''
        r['Acct Type'] = 'Acct'
        r['Notif Month'] = '-'.join([r['Effective Year'], month_num[r['Effective Month']]])
        r['District'] = enroldistrict.get(r['Enrollment #'], '')
        if r['Type'] == 'Win': r['COCP Win'] = r['Master #']
        else: r['COCP Loss'] = r['Master #']

    elif datatype == 'NNEAEnrol':
        r['Acct Type'] = 'Enrollment'
        r['Notif Month'] = dparser.parse(r['Contract Create Date']).date().strftime("%Y-%m")
        if r['Contract Category'] == 'New Contract': r['NNEA'] = r['Contract Number']
        else: r['Renewal'] = r['Contract Number']
    elif datatype == 'NNEAAcct':
        r['Acct Type'] = 'Acct'
        r['Contract Program Name'] = ''
        r['Notif Month'] = dparser.parse(r['Contract Create Date']).date().strftime("%Y-%m")
        if r['Contract Category'] == 'New Contract': r['NNEA'] = r['Master Number']
        else: r['Renewal'] = r['Master Number']
    r['Rep'] = r['Rep'].strip()
    r['Customer'] = r['Customer'].upper()

    # Correct Rep Type / Branch
    try:
        r['Rep Type'] = titleOBTSR[emp_stb[r['Rep']][1]]
    except KeyError:
        if 'Coverage' in r:
            r['Rep Type'] = r['Coverage']
            if r['Rep Type'] == 'TB':
                r['Rep Type'] = 'TSR'
    try:
        r['Branch'] = emp_stb[r['Rep']][0]
    except KeyError:
        pass
    return r

############### main ###############

def main():
    t0 = time.clock()
    header = get_header()
    split_enrolprogramdivloc()

    # Analyze NNEA/COCP -> Output
    with open(output, 'wb') as o:
        writer = csv.writer(o)
        ow = csv.DictWriter(o, fieldnames=header)
        writer.writerows([header])

        scan_COCP(ow)
        scan_NNEA(ow)

    t1 = time.clock()
    print 'Process completed! Duration:', t1-t0
    return

if __name__ == '__main__':
    main()

