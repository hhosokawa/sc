from aux_reader import *
from collections import defaultdict
import csv
import time
from pprint import pprint

""" io """

bi          = 'i/dmoo_growth/bi.csv'
esa_booking = 'i/dmoo_growth/esa_booking.csv'
ref         = 'i/dmoo_growth/ref.csv'
output      = 'o/dmoo_growth.csv'

current_month       = 5
current_year        = 2014
div_dict            = {'100': 'Canada',
                       '200': 'United States'}
esa_booking_rows    = []
esa_ref             = set()
header              = []

############### utils ###############

# Obtain Headers from BI
def get_header():
    with open(bi) as i3:
        return sorted(csv.DictReader(i3).fieldnames)

# Scan Referrals for ESA Order #
def scan_ref():
    with open(ref) as i1:
        for r in csv.DictReader(i1):
            if r['Referral Source'] in ['ESA 3.0', 'ESA 4.0']:
                esa_ref.add(r['Referral Number'])
    print 'scan_ref() complete.'

# Absorb ESA Booking Rows
def scan_esa_booking():
    with open(esa_booking) as i2:
        for r in csv.DictReader(i2):
            d = defaultdict(str)
            d['Calendar Month'] = current_month
            d['Calendar Year'] = current_year
            d['District'] = r['DISTRICT_NAME']
            d['Division'] = div_dict.get(r['DIVISION'], 'N/A')
            d['Region'] = r['REGION_NAME']
            d['Solution Group'] = 'ESA ESTIMATE'
            d['Super Category'] = 'Microsoft'
            d['Virtually Adjusted GP'] = r['GP_FORECAST']
            d['Virtually Adjusted Imputed Revenue'] = r['GP_FORECAST']
            d['Virtually Adjusted Revenue'] = r['GP_FORECAST']
            esa_booking_rows.append(d)
    print 'scan_esa_booking() complete.'

# Clean BI and write_csv()
def write_csv():
    with open(output, 'wb') as o:
        writer = csv.writer(o)
        header = get_header()
        ow = csv.DictWriter(o, fieldnames=header)
        writer.writerows([header])
        with open(bi) as i3:

            # Write BI, all ESA order # -> 1 month back
            for r in csv.DictReader(i3):
                if (r['Calendar Year'] == '2014'
                and r['Order Number'] in esa_ref
                and int(r['Calendar Month']) > 1):
                    r['Calendar Month'] = int(r['Calendar Month']) - 1
                ow.writerow(r)
            print 'write_csv() - BI data complete.'

            # Write ESA Estimate Data
            for r in esa_booking_rows:
                ow.writerow(dict(r))
            print 'write_csv() - ESA Estimate data complete.'
    print 'write_csv() complete.'

############### ms_sales_main() ###############

if __name__ == '__main__':
    t0 = time.clock()
    scan_ref()
    scan_esa_booking()
    write_csv()
    t1 = time.clock()
    print 'ms_sales_main() completed! Duration:', t1-t0
