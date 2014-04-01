import csv
import time
import pprint
from aux_reader import *
import dateutil.parser as dparser
from collections import OrderedDict
from datetime import timedelta, date

""" io """
output = 'o/ms_incentive.csv'
input0 = 'i/ms_incentive/ms_sales.csv'

rows = []

############## utils ###############

def get_header():
    header = set()
    with open(input0) as i0: header.update(csv.DictReader(i0).fieldnames)
    header = tuple(sorted(header, key=lambda item: item[0]))
    return header

def determine_new_product(r):
    # Determine Admissible New Product Category
    if r['Category C'] in ['EA Add On', 'EA Add-On',
                           'EA True-Up', 'EA True Up',
                           'EA New']:
        return True
    elif r['Category B'] == 'MS OSA':
        return True
    elif r['Category A'] in ['Open', 'Select']:
        return True
    else:
        return False

def correct_month(r):
    # If Year < 2013: New Month = Month - 1
    year = int(r['Calendar Year'])
    month = int(r['Calendar Month'])

    if (r['Calendar Year'] == '2013' and
        r['Category A'] not in ['Open', 'Select']):
        if r['Calendar Month'] == '1':
            pass
        else:
            r['Calendar Month'] = month - 1
            return r
    else:
        return r

def district_tsr_switch(r):
    # Make TSR Reps roll up to OB/TSR Manager
    if r['OB or TSR'] == 'TSR':
        r['District'] = r['OB/TSR Manager'].title()
    return r

def scan_ms_sales():
    with open(input0) as i0:
        for r in csv.DictReader(i0):
            if determine_new_product(r):
                correct_month(r)
                district_tsr_switch(r)
                rows.append(r)
    print 'scan_ms_sales() complete.'

def write_csv():
    with open(output, 'wb') as o:
        writer = csv.writer(o)
        ow = csv.DictWriter(o, fieldnames=header)
        writer.writerows([header])
        for r in rows:
            ow.writerow(r)
    print 'write_csv() complete.'

############## sb_renewals_main() ###############

if __name__ == '__main__':
    t0 = time.clock()
    header = get_header()
    scan_ms_sales()
    write_csv()
    t1 = time.clock()
    print 'ms_incentive_main() completed. Duration:', t1-t0
