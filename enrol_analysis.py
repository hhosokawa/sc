import csv
import time
from aux_reader import *
from datetime import datetime
import dateutil.parser as dparser
from dateutil.relativedelta import relativedelta

""" io """
enrol_path = 'i/enrol_analysis/enrol test.csv'
enrol_out = 'o/enrol_analysis.csv'

""" pictionary """
enrols = {}
rows = []
divregion = csv_dic('auxiliary\\div-region.csv')
divdistrict = csv_dic('auxiliary\\div-district.csv')

############### aux ###############

def diff_month(d1, d2):
    return (d1.year - d2.year)*12 + d1.month - d2.month

def get_region_district(r):
    divloc = str(r['Master Division']) + str(r['Master Divloc'])
    region = divregion.get(divloc, '')
    district = divdistrict.get(divloc, '')
    return region, district

############### utils ###############

def scan_enrol():
    with open(enrol_path) as i_enrol:
        for r in csv.DictReader(i_enrol):
            if r['Contract Number'] not in enrols:
                enrols[r['Contract Number']] = r
            else:
                date_x = dparser.parse(enrols[r['Contract Number']]
                                       ['Contract Create Date'])
                date_y = dparser.parse(r['Contract Create Date'])
                if date_y > date_x:
                    enrols[r['Contract Number']] = r
    print 'scan_enrol() complete.'

def extract_valid_months():
    for e in enrols:
        contract_status = enrols[e]['Contract Status']
        end_date =  dparser.parse(enrols[e]['Contract End Date'])
        start_date = dparser.parse(enrols[e]['Contract Start Date'])
        relative_months = diff_month(end_date, start_date)
        if contract_status == 'Active':
            period = start_date
            for month in range(relative_months):
                number = enrols[e]['Contract Number']
                category =  enrols[e]['Contract Category']
                mm_name = enrols[e]['Master-Master Name']
                rep = enrols[e]['Master OB Rep Name']
                lic_program = enrols[e]['Licensing ProgramName']
                region, district = get_region_district(enrols[e])
                year = period.strftime('%Y')
                month = period.strftime('%m')
                row = (number, year, month, region, district, 
                       category, rep, lic_program)
                rows.append(row)
                period = period + relativedelta(months=1)
                if period >= datetime.today():
                    break
    enrols.clear()
    print 'extract_valid_months() complete.'

def write_out():
    header = ['Enrol #', 'Year', 'Period', 'Region', 'District',
              'Category', 'Rep', 'Lic Program']
    with open(enrol_out, 'wb') as o:
        writer = csv.writer(o)
        writer.writerow(header)
        for row in rows:
            writer.writerow(row)
    print 'write_out() complete.'

############### enrol_main() ###############

def enrol_main():
    t0 = time.clock()
    scan_enrol()
    extract_valid_months()
    write_out()
    t1 = time.clock()
    print 'enrol_main() completed. Duration: ', t1-t0

if __name__ == '__main__':
    enrol_main()
