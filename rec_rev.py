import re
import csv
import time
from aux_reader import *

output = 'o/rec_rev.csv'

bi = 'i/rec_rev/rec_rev_best_p10.csv'
lic_valid_per_items = csv_dic('i/rec_rev/lic_valid_period_items.csv', 2)
rev_rec_flag = ['SWBNDL', 'HWMTN', 'SWMTN', 'CLOUDN', 'SWSUB', 'CLOUD', 
                'TRAIN']

############### Utils ###############

def get_headers():
    with open(bi) as bi1:
        headers = csv.DictReader(bi1).fieldnames
    headers.append('Item Class')
    return sorted(headers)

def write_headers(o, headers):
    writer = csv.writer(o)
    writer.writerows([headers])

def loop_row(i1, ow):
    i1r = csv.DictReader(i1)
    for r in i1r:
        gp = float(r['GP (includes Freight)'])
        rev = float(r['Net Revenue'])

        if (r['Item Number'] in lic_valid_per_items or
            r['Sale or Referral'] == 'Referral' or
            r['Super Category @ Order Date'] == 'Managed Services' or
            r['Revenue Recognition ID'] in rev_rec_flag):
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
