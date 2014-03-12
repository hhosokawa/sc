import csv
import time
import pprint
from aux_reader import *
import dateutil.parser as dparser
from collections import OrderedDict
from datetime import timedelta, date

""" io """
data = {}
fb_data = tree()
current_year = 2014
input0 = 'i/sb_renewals/contract_repo.csv'
input1 = 'i/sb_renewals/futurebillings.csv'
output = 'P:/_HHOS/Microsoft/SB Renewals - %s.csv' % (time.strftime("%Y-%m-%d"))

divregion = csv_dic('auxiliary/div-region.csv')
divdistrict = csv_dic('auxiliary/div-district.csv')

############## utils ###############

def scan_future_billing():
    with open(input1) as i1:
        for r in csv.DictReader(i1):
            contract = r['Agreement Number']
            sb_date = r['Scheduled Bill Date']
            if r['Extended Amount']:
                imputed_rev = float(r['Extended Amount'])

                if ((contract not in fb_data)
                or (contract in fb_data and sb_date not in fb_data[contract])):
                    fb_data[contract][sb_date] = imputed_rev
                else:
                    fb_data[contract][sb_date] += imputed_rev
    print 'scan_future_billing() completed.'

def calc_gp(r):
    # Assign Contract Units
    if r['Contract Units']:
        units = int(r['Contract Units'])
    else: units = -1

    # Assign Imputed Revenue
    if r['Contract Payment Amount']:
        imputed_rev = float(r['Contract Payment Amount'])
    elif r['Contract Number'] in fb_data:
        ordered = OrderedDict(sorted(fb_data[r['Contract Number']].items(),
                              key = lambda t: t[0]))
        imputed_rev = ordered.items()[0][1]
    else:
        imputed_rev = 0

    # GP Fee Matrix
    start_date = dparser.parse(r['Contract Payment Date'])
    if dparser.parse('2013-09-30') < start_date:

        # ESA 3.0 date < 2013-09-30
        if 0 <= units <= 749:           percent = 0.06
        elif 750 <= units <= 2399:      percent = 0.06
        elif 2400 <= units <= 5999:     percent = 0.055
        elif 6000 <= units <= 14999:    percent = 0.033
        elif units >= 15000:            percent = 0.03
        else:                           percent = 0.048
    else:

        # ESA 4.0 - Renewal
        if r['Type'] == 'Renewal':
            if 0 <= units <= 749:           percent = 0.085
            elif 750 <= units <= 2399:      percent = 0.0488
            elif 2400 <= units <= 5999:     percent = 0.0375
            elif 6000 <= units <= 14999:    percent = 0.0113
            elif units >= 15000:            percent = 0.0083
            else:                           percent = 0.0382

        # ESA 4.0 - Scheduled Billings
        else:
            if 0 <= units <= 749:           percent = 0.025
            elif 750 <= units <= 2399:      percent = 0.01125
            elif 2400 <= units <= 5999:     percent = 0.01
            elif 6000 <= units <= 14999:    percent = 0.00625
            elif units >= 15000:            percent = 0.0045
            else:                           percent = 0.0114

    return imputed_rev * percent

def get_region_district(r):
    if r['Master Division'] and r['Master Divloc']:
        divloc = r['Master Division'] + r['Master Divloc']
        return divregion.get(divloc), divdistrict.get(divloc)
    else: return 'N/A', 'N/A'

def scan_contract_repo():
    with open(input0) as i0:
        for r in csv.DictReader(i0):
            level = r['Contract Level']
            status = r['Contract Status']
            contract = r['Contract Number']
            end = dparser.parse('2014-12-31')
            start = dparser.parse('2014-01-01')
            end_date = dparser.parse(r['Contract End Date'])
            pay_date = dparser.parse(r['Contract Payment Date'])

            if status == 'Active' and contract[0] != 'S':
                if contract in data:
                    pay_date2 = dparser.parse(data[contract]
                                ['Contract Payment Date'])
                    if pay_date > pay_date2:
                        pass
                else:
                    if start <= end_date <= end:
                        r['Type'] = 'Renewal'
                        r['Fiscal Year'] = end_date.strftime('%Y')
                        r['Fiscal Month'] = end_date.strftime('%m')
                    elif start <= pay_date <= end:
                        r['Type'] = 'Scheduled Billing'
                        r['Fiscal Year'] = pay_date.strftime('%Y')
                        r['Fiscal Month'] = pay_date.strftime('%m')
                    else:
                        continue
                    r['GP'] = calc_gp(r)
                    r['Imputed Rev'] = r['Contract Payment Amount']
                    r['Region'], r['District'] = get_region_district(r)

                    data[contract] = r

    print 'scan_contract_repo() completed.'

def add_renewal_trueup():
    for contract in data:
        sb_renewal = data[contract]['Type']
        end_date = dparser.parse(data[contract]['Contract End Date']).date()
        start_date = dparser.parse(data[contract]['Contract Start Date']).date()
        true_up_date = start_date - timedelta(days = 30)

        # True-Up Date (Feb 29 Fix)
        try:
            true_up_date = true_up_date.replace(year = current_year)
        except ValueError:
            true_up_date = true_up_date.replace(day = 28)
            true_up_date = true_up_date.replace(year = current_year)
        data[contract]['True Up Date'] = true_up_date.strftime('%Y-%m-%d')

        if sb_renewal == 'Renewal':
            renewal_date = end_date + timedelta(days = 1)
            data[contract]['Renewal Date'] =renewal_date.strftime('%Y-%m-%d')

def write_csv():
    headers = ['Contract Number', 'Contract Start Date',
               'Contract End Date','Contract Program Name',
               'Contract Units', 'Master Number', 'Master Name',
               'Imputed Rev', 'GP', 'Region', 'District', 'Master OB Rep Name',
               'Fiscal Year', 'Fiscal Month', 'Type', 'True Up Date',
               'Renewal Date']

    with open(output, 'wb') as o0:
        o0w = csv.DictWriter(o0, delimiter=',',
                             fieldnames=headers, extrasaction='ignore')
        o0w.writerow(dict((fn, fn) for fn in headers))
        for contract in data:
            o0w.writerow(data[contract])
    print 'write_csv() completed.'

############## sb_renewals_main() ###############

if __name__ == '__main__':
    t0 = time.clock()
    scan_future_billing()
    scan_contract_repo()
    add_renewal_trueup()
    write_csv()
    t1 = time.clock()
    print 'sb_renewals_main() completed. Duration:', t1-t0
