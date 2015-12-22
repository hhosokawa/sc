import csv
import time

""" io """
INPUT1 = 'i/ms_sales/bi.csv'
OUTPUT = 'o/ms_sales.csv'
enrollment_ids = set()

# Converts Csv -> Dictionary
def csv_dic(filename):
    reader = csv.reader(open(filename, "rb"))
    my_dict = dict((k, v) for k, v in reader)
    return my_dict

# Pictionary
referral_enrollment = csv_dic('i/ms_sales/referral_enrollment.csv')
super_prodtype = csv_dic('i/ms_sales/super_prodtype.csv')

############### utils ###############

# Obtain Headers
def get_header():
    header = set()
    with open(INPUT1) as i1: header.update(csv.DictReader(i1).fieldnames)
    new_fields = set(['Category A', 'Category B', 'Category C', 
                      'Level', 'Enrollment Number', 'Unique Enrollment ID'])
    header = new_fields | header
    try: header.remove('')
    except KeyError: pass
    header = tuple(sorted(header, key=lambda item: item[0]))
    return header

#______ BI data _______
# Add Category to Row
def add_cat(r):
    if r['Sale or Referral'] == 'Sale':
        r['Category A'] = super_prodtype.get(r['ProdType Ven Program'], 
                                             'EA Indirect and Other')
        r['Category B'] = r['ProdType Ven Program']
    elif r['Sale or Referral'] == 'Referral':
        r['Category A'] = r['Referral Source']
        r['Category B'] = r['Referral Revenue Type']
        r['Enrollment Number'] = referral_enrollment.get(r['Order Number'],'')
        enrollment_id = r['Calendar Year'] + r['Calendar Month'] + r['Category B'] + r['Enrollment Number']
        # Include Unique Enrollment IDs for 
        if enrollment_id not in enrollment_ids:
            enrollment_ids.add(enrollment_id)
            r['Unique Enrollment ID'] = enrollment_id
    return r

# Insert Customer Level (A1, A2, B, etc.)
def add_customer_level(r):
    try:
        pc_count = int(r['Master PC Count'])
    except ValueError:
        pc_count = 0
    if 0 <= pc_count <= 749:            r['Level'] = 'A1'
    elif 750 <= pc_count <= 2399:       r['Level'] = 'A2'
    elif 2400 <= pc_count <= 5999:      r['Level'] = 'B'
    elif 6000 <= pc_count <= 15000:     r['Level'] = 'C'
    elif pc_count > 15000:              r['Level'] = 'D'
    else:                               r['Level'] = 'N/A'
    return r

def write_csv():
    with open(OUTPUT, 'wb') as o:
        writer = csv.writer(o)
        ow = csv.DictWriter(o, fieldnames=header)
        writer.writerows([header])

        # BI Data + Product Catagory Data
        with open(INPUT1) as i1:
            for r in csv.DictReader(i1):
                add_cat(r)
                add_customer_level(r)
                ow.writerow(r)

############### ms_sales_main() ###############

if __name__ == '__main__':
    t0 = time.clock()
    header = get_header()
    write_csv()
    t1 = time.clock()
    print 'ms_sales_main() completed! Duration:', t1-t0