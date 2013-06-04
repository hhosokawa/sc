import csv
import time
from aux_reader import *
import dateutil.parser as dparser
from datetime import datetime, timedelta, date
from collections import defaultdict

output = 'o/NNEA COCP Summary - 2013-06-03.csv'
input1 = 'i/ms_cocp/COCP Summary to 2013-05-15.csv'
input2 = 'i/ms_cocp/NNEA Summary to 2013-06-03.csv'
enrolprogram = csv_dic('auxiliary\\enrol-program.csv')
emp_stb = csv_dic('auxiliary\\employee-super_title_branch.csv', 3)

# Pictionary Jars
divregion = csv_dic('auxiliary\\div-region.csv')
titleOBTSR = csv_dic('auxiliary\\title-OB_TSR.csv')
enrolset, accttypeset, masterset = set(), set(), set()
month_num = {'Jan': '01', 'Feb': '02', 'Mar': '03',
             'Apr': '04', 'May': '05', 'Jun': '06',
             'Jul': '07', 'Aug': '08', 'Sep': '09',
             'Oct': '10', 'Nov': '11', 'Dec': '12'}

#################################################################################
## Function Definition

# Adds New Headers
def header_add(header):
    newfields = ['NNEA', 'Renewal', 'COCP Win',
                 'COCP Loss', 'Rep Type', 'Acct Type', 'Notif Month']
    for newfield in newfields: header.add(newfield)
    return header

def clean(r, datatype):

    if datatype == 'COCPEnrol':
        r['Notification Month'] = r['Notification Month'].strip()
        r['Acct Type'] = 'Enrollment'
        r['Contract Program Name'] = enrolprogram.get(r['Enrollment #'], '')
        r['Notif Month'] = '-'.join([r['Notif Year'], month_num[r['Notification Month']]])
        if r['Type'] == 'Win': r['COCP Win'] = r['Enrollment #']
        else: r['COCP Loss'] = r['Enrollment #']
    elif datatype == 'COCPAcct':
        r['Notification Month'] = r['Notification Month'].strip()
        r['Contract Program Name'] = ''
        r['Acct Type'] = 'Acct'
        r['Notif Month'] = '-'.join([r['Notif Year'], month_num[r['Notification Month']]])
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
    r['OB Rep'] = r['OB Rep'].strip()
    r['Customer'] = r['Customer'].upper()

    # Correct Rep Type / Branch
    try:
        r['Rep Type'] = titleOBTSR[emp_stb[r['OB Rep']][1]]
    except KeyError:
        if 'Coverage' in r:
            r['Rep Type'] = r['Coverage']

        if r['Rep Type'] == 'TB':
            r['Rep Type'] = 'TSR'
    try:
        r['Branch'] = emp_stb[r['OB Rep']][0]
    except KeyError:
        pass
    return r

#################################################################################
## Main
def main():
    t0 = time.clock()

    # All input headers -> output header
    header = set()
    with open(input1) as i1: header.update(csv.DictReader(i1).fieldnames)
    with open(input2) as i2: header.update(csv.DictReader(i2).fieldnames)
    header = header_add(header)
    header = tuple(header)

    # Analyze Input 1-2 -> Output
    with open(output, 'wb') as o:
        writer = csv.writer(o)
        ow = csv.DictWriter(o, fieldnames=header)
        writer.writerows([header])
        with open(input1) as i1:
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

        with open(input2) as i2:
            i2r = csv.DictReader(i2)

            # NNEA Enrollment Cleaning
            for r in i2r:
                r['OB Rep'] = r['Master OB Rep Name']
                r['Master #'] = r['Master Number']
                r['Customer'] = r['Master Name']
                masterdivloc = r['Master Division'] + r['Master Divloc']
                r['Region'] = divregion[masterdivloc]
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

    t1 = time.clock()
    return 'Process completed! Duration:', t1-t0

print main()
