from aux_reader import *
import csv
from dateutil.parser import parse
from collections import defaultdict
from itertools import product
from pprint import pprint
import time

# io
input_ms_contracts = 'i/ms_dashboard/ms_contracts.csv'
input_ms_sales = 'i/ms_dashboard/ms_sales_unique_customer.csv'

months = range(3,13,3)
quarters = {1:1, 2:1, 3:1, 4:2, 5:2, 6:2,
            7:3, 8:3, 9:3, 10:4, 11:4, 12:4}
years = range(2011, 2015)

############### utils ###############

# Unique Contracts
def scan_ms_contracts():
    U = tree() # Initialize Matrix
    for r in csv.DictReader(open(input_ms_contracts)):
        start = parse(r['Contract Start Date'])
        end = parse(r['Contract End Date'])
        for y, m in product(years, months):
            program = r['Licensing ProgramName']
            q = quarters[m]
            target_date = parse(str(y)+'-'+str(m)+'-'+'01')

            if start <= target_date <= end:
                if U[y][q][program]:
                    U[y][q][program].add(r['Contract Number'])
                else:
                    U[y][q][program] = set([r['Contract Number']])

    # Print Contract Matrix
    print 'Unique Contracts'
    print '================'
    for y, m in product(years, months):
        q = quarters[m]
        for program in U[y][q].keys():
            print y,q,program,len(U[y][q][program])

# Unique Customers
def scan_ms_sales():
    U = tree() # Initialize Matrix
    for r in csv.DictReader(open(input_ms_sales)):
        master_class = r['Master Classification']
        y = int(r['Calendar Year'])
        q = int(r['Fiscal Quarter'])
        if U[y][q][master_class]:
            U[y][q][master_class].add(r['Unique Master Master Name'])
        else:
            U[y][q][master_class] = set([r['Unique Master Master Name']])

    # Print Customer Matrix
    print '\n','Unique Customers'
    print '================'
    for y, q in product(years, range(1,5)):
        for master_class in U[y][q].keys():
            print y,q,master_class,len(U[y][q][master_class])

############### main ###############

if __name__ == '__main__':
    t0 = time.clock()
    #scan_ms_contracts()
    scan_ms_sales()
    t1 = time.clock()
    print 'ms_dashboard complete.', t1-t0