import re
import csv
import time
from aux_reader import *

output = 'o/rec_rev.csv'
bi = 'i/rec_rev/2013 ytd.csv'
lic_valid_per_items = csv_dic('i/rec_rev/lic_valid_period_items.csv', 2)

############### Utils ###############

def get_headers():
    with open(bi) as bi1:
        headers = csv.DictReader(bi1).fieldnames
    headers.append('Item Class')
    return sorted(headers)

def write_headers(o, headers):
    writer = csv.writer(o)
    writer.writerows([headers])

def str_float(s):
    s = s.replace('$', '')
    s = s.replace(',', '')
    if '(' in s:
        s = s.replace('(' , '-')
        s = s.replace(')', '')
    return float(s)

def loop_row(i1, ow):
    i1r = csv.DictReader(i1)
    for r in i1r:
        gp = str_float(r['Virtually Adjusted GP'])
        rev = str_float(r['Virtually Adjusted Revenue'])

        if (r['Item Number'] in lic_valid_per_items or
            r['Sale or Referral'] == 'Referral' or
            r['Solution Group'] == 'SERVICES'):
            r['Item Class'] = 'Recurring Revenue'
        elif (gp > 1000 and rev > 10000):
            r['Item Class'] = 'Project'
        else:
            r['Item Class'] = 'Run Rate'
   
        ow.writerow(r)

############### Main ###############

def rec_rev_main():
    t0 = time.clock()
    headers = get_headers()

    with open(output, 'wb') as o:
        ow = csv.DictWriter(o, fieldnames=headers)
        write_headers(o, headers)
      
        with open(bi, 'rb') as i1:
            loop_row(i1, ow)

    t1 = time.clock()
    print 'Process completed! Duration:', t1-t0
    return 

if __name__ == '__main__':
    rec_rev_main()
