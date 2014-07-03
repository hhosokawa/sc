from aux_reader import *
import csv
import time

input0 = 'i/rec_rev/bi.csv'
output = 'o/rec_rev.csv'

rev_rec_flag = ['SWBNDL', 'HWMTN', 'SWMTN', 'CLOUDN',
                'SWSUB', 'CLOUD', 'TRAIN']

############### Utils ###############

def get_headers():
    with open(input0) as i0:
        headers = csv.DictReader(i0).fieldnames
    headers.append('Item Class')
    return sorted(headers)

def loop_row(i0, ow):
    i0r = csv.DictReader(i0)
    for r in i0r:
        gp = float(r['GP (includes Freight)'])
        rev = float(r['Revenue (includes Freight)'])

        if (r['Sale or Referral'] == 'Referral' or
            r['Revenue Recognition ID'] == 'MSRV' or
            r['Super Category'] == 'Managed Services'):
            r['Item Class'] = 'Recurring Revenue'
        elif r['Revenue Recognition ID'] in rev_rec_flag:
            r['Item Class'] = 'Potentially Recurring Revenue'
        elif (gp > 1000 and rev > 10000):
            r['Item Class'] = 'Project'
        else:
            r['Item Class'] = 'Run Rate'
        ow.writerow(r)

def write_csv(headers):
    with open(output, 'wb') as o:
        ow = csv.DictWriter(o, fieldnames=headers)
        writer = csv.writer(o)
        writer.writerows([headers])
        with open(input0, 'rb') as i0:
            loop_row(i0, ow)

############### Main ###############

if __name__ == '__main__':
    t0 = time.clock()
    headers = get_headers()
    write_csv(headers)
    t1 = time.clock()
    print 'Process completed! Duration:', t1-t0
