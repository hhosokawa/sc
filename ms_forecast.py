from aux_reader import *
import csv
from datetime import datetime, timedelta
from dateutil.parser import parse
from pprint import pprint
import time

""" io """
output = 'o/ms_forecast.csv'
input1 = 'i/ms_forecast/toad_data.csv'

district_dict = csv_dic('auxiliary/div-district.csv')
region_dict = csv_dic('auxiliary/div-region.csv')

rows = []

############### utils ###############

def scan_toad_data():
    with open(input1) as i1:
        for r in csv.DictReader(i1):
            # Clean Data
            r['Period'] = r['Period'][1:]       # Remove P1, P2, Pn
            if r['Manage&Deploy'] == 'Yes':     # Imputed Revenue = 0
                r['TOTALREVENUE'] = 0
        
            

  
            rows.append(r)
    print 'completed scan_toad_data()'

def write_csv():
    headers = ['BRANCH_NAME', 'CONTRACTID', 'DISTRICT_NAME', 'DIVISION', 
               'GP_FORECAST', 'Manage&Deploy', 'Manual Add', 'NAME', 'OB_NAME',
               'OB_TSR', 'OB_TSR_FLAG', 'PAYMENTTYPE_POTYPE', 'Period', 
               'REGION_NAME', 'TOTALREVENUE']
    
    with open(output, 'wb') as o1:
        dw = csv.DictWriter(o1, fieldnames = sorted(headers))
        dw.writerow(dict((h,h) for h in headers))
        for r in rows:
            dw.writerow(r)
    print 'completed write_csv()'

############### ms_forecast_main() ###############

if __name__ == '__main__':
    t0 = time.clock()
    scan_toad_data()
    write_csv()
    t1 = time.clock()
    print 'ms_forecast_main() completed! Duration:', t1-t0
