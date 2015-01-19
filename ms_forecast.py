from aux_reader import *
import csv
from datetime import datetime, timedelta
from dateutil.parser import parse
from pprint import pprint
import time

""" io """
output = 'o/ms_forecast.csv'
input1 = 'i/ms_forecast/repman_contract_repo.csv'

district_dict = csv_dic('auxiliary/div-district.csv')
region_dict = csv_dic('auxiliary/div-region.csv')
tsr_reps_dict = csv_dic('i/ms_forecast/bi_tsr_reps.csv')

rows = []
headers = []

############### utils ###############

def scan_contract_repo():
    with open(input1) as i1:
        for r in csv.DictReader(i1):
            # Populate Region / District
            div_loc = r['Master Division'] + r['Master Divloc']
            r['Region'] = region_dict.get(div_loc, '')
            r['District'] = district_dict.get(div_loc, '')
            
            # Populate Contract Level
            if r['Contract Units']: units = int(r['Contract Units'])
            else: units = 0
            if 0 <= units <= 749: r['Level'] = 'A1'
            elif 750 <= units <= 2399: r['Level'] = 'A2'
            elif 2400 <= units <= 5999: r['Level'] = 'B'
            elif 6000 <= units <= 15000: r['Level'] = 'C'
            elif units > 15000: r['Level'] = 'D'

            # Populate OB / TSR
            if r['Master OB Rep'] in tsr_reps_dict:
                r['OB / TSR'] = 'TSR'
            else:
                r['OB / TSR'] = 'OB'
   
            # Populate Renewal / Scheduled Billing
            end_date = parse(r['Contract End Date'])
            renewal_date = datetime(2015, 12, 30)
            if end_date <= renewal_date:
                r['Renewal / Scheduled Billing'] = 'Renewal'
            else:
                r['Renewal / Scheduled Billing'] = 'Scheduled Billing'
            
            # Populate Year, Month
            start_date = parse(r['Contract Start Date'])
            r['Year'] = '2015'
            r['Month'] = start_date.month

            # Populate True Up Date, Renewal Date
            if start_date.month != 1:
                true_up_month = start_date.month - 1
            else:
                true_up_month = 1
            r['True Up Date'] = datetime(2015, true_up_month, 1).strftime('%Y-%m-%d')
            if r['Renewal / Scheduled Billing'] == 'Renewal':
                renewal_date = end_date + timedelta(days=1)
                r['Renewal Date'] = renewal_date.strftime('%Y-%m-%d')
                global headers
                if not headers: headers = r.keys()
  
            rows.append(r)
    print 'completed scan_contract_repo()'

def write_csv():
    with open(output, 'wb') as o1:
        dw = csv.DictWriter(o1, fieldnames = sorted(headers))
        dw.writerow(dict((h,h) for h in headers))
        for r in rows:
            dw.writerow(r)
    print 'completed write_csv()'

############### ms_forecast_main() ###############

if __name__ == '__main__':
    t0 = time.clock()
    scan_contract_repo()
    write_csv()
    t1 = time.clock()
    print 'ms_forecast_main() completed! Duration:', t1-t0
