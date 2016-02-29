import csv
from pprint import pprint
import time

################ Input / Output ################
INPUT = 'i/ms_sales/bi.csv'
OUTPUT = 'P:/_HHOS/Microsoft/data/ms_sales_data.csv'
enrollment_ids = set()

def csv_dic(filename):
    # Converts .csv -> Dictionary
    reader = csv.reader(open(filename, "rb"))
    my_dict = dict((k, v) for k, v in reader)
    return my_dict
    
# Auxiliary Dictionaries
contract_number_units = csv_dic('i/ms_sales/contract_number_units.csv')
super_prodtype = csv_dic('i/ms_sales/super_prodtype.csv')

############### Utils ###############
def get_header():
    # Obtain Headers
    header = set()
    with open(INPUT) as i1: header.update(csv.DictReader(i1).fieldnames)
    new_fields = set(['Category A', 'Category B', 'Level', 'Unique Enrollment ID',
                      'Contract Units', 'Sub 500'])
    header = new_fields | header
    try: header.remove('')
    except KeyError: pass
    header = tuple(sorted(header, key=lambda item: item[0]))
    return header

def add_cat(r):
    # Add Category to Row
    if r['Sale or Referral'] == 'Sale':
        r['Category A'] = super_prodtype.get(r['ProdType Ven Program'], 
                                             'EA Indirect and Other')
        r['Category B'] = r['ProdType Ven Program']
    elif r['Sale or Referral'] == 'Referral':
        r['Category A'] = r['Referral Source']
        r['Category B'] = r['Referral Revenue Type']
        
        # Add Desktop Count
        r['Contract Units'] = contract_number_units.get(r['Referral Enrollment ID'], 0)
        if r['Contract Units']: 
            r['Contract Units'] = int(r['Contract Units'])
        
        # Include Unique Enrollment IDs
        enrollment_id = r['Calendar Year'] + r['Calendar Month'] + r['Category B'] + r['Referral Enrollment ID']
        if enrollment_id not in enrollment_ids:
            enrollment_ids.add(enrollment_id)
            r['Unique Enrollment ID'] = enrollment_id
    return r

def add_customer_level(r):
    # Determine Customer Level (A1, A2, B, etc.)
    if r['Sale or Referral'] == 'Sale':
        r['Level'] = 'N/A'
        r['Sub 500'] = 'FALSE'
        return r
    elif r['Sale or Referral'] == 'Referral':
        seat_count = r['Contract Units']
        if 1 <= seat_count <= 250:          
            r['Level'] = 'A1 EA 0-250'
            r['Sub 500'] = 'TRUE'
        elif 251 <= seat_count <= 500:      
            r['Level'] = 'A1 EA 251-500'
            r['Sub 500'] = 'TRUE'
        elif 501 <= seat_count <= 749:      
            r['Level'] = 'A1 EA 501-749'
            r['Sub 500'] = 'FALSE'
        elif 750 <= seat_count <= 2399:     
            r['Level'] = 'A2'
            r['Sub 500'] = 'FALSE'
        elif 2400 <= seat_count <= 5999:    
            r['Level'] = 'B'
            r['Sub 500'] = 'FALSE'
        elif 6000 <= seat_count <= 15000:   
            r['Level'] = 'C'
            r['Sub 500'] = 'FALSE'
        elif (seat_count > 15000) and seat_count:            
            r['Level'] = 'D'
            r['Sub 500'] = 'FALSE'
        else:                               
            r['Level'] = 'N/A'
            r['Sub 500'] = 'FALSE'
        return r
        
def write_csv():
    with open(OUTPUT, 'wb') as o:
        writer = csv.writer(o)
        ow = csv.DictWriter(o, fieldnames=header)
        writer.writerows([header])

        # BI Data + Product Category Data
        with open(INPUT) as i1:
            for r in csv.DictReader(i1):
                add_cat(r)
                add_customer_level(r)
                ow.writerow(r)

############### Main ###############
if __name__ == '__main__':
    t0 = time.clock()
    header = get_header()
    write_csv()
    t1 = time.clock()
    print 'ms_sales_main() completed! Duration:', t1-t0