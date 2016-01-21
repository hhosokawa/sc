import csv
from itertools import product
import pandas as pd
import time

# Converts csv -> Dictionary
def csv_dic(filename):
    reader = csv.reader(open(filename, "rb"))
    my_dict = dict((k, v) for k, v in reader)
    return my_dict 

############### Data Sources ###############
annual_forecast = pd.read_csv('i/rebate_mdf_upload/annual_forecast.csv')
div_100_branch_split = pd.read_csv('i/rebate_mdf_upload/div_100_branch_split.csv')
div_200_branch_split = pd.read_csv('i/rebate_mdf_upload/div_200_branch_split.csv')
gl_split = pd.read_csv('i/rebate_mdf_upload/gl_split.csv')
monthly_seasonality = pd.read_csv('i/rebate_mdf_upload/monthly_seasonality.csv')
usd_cad_fx = 0.76   # USD->CAD FX
OUTPUT_CSV = 'o/rebate_mdf_upload_2016.csv'

# Data Source Headers
div_100_branch_split_headers = div_100_branch_split.columns.tolist()
div_200_branch_split_headers = div_200_branch_split.columns.tolist()
gl_split_headers = gl_split.columns.tolist()
months = monthly_seasonality.columns.tolist()
rows = annual_forecast.index

# Pictionary
categories = csv_dic('i/rebate_mdf_upload/auxiliary/categories.csv')
gl_parents = csv_dic('i/rebate_mdf_upload/auxiliary/gl_parent.csv')
quarters = {'1':'1', '2':'1', '3':'1', 
            '4':'2', '5':'2', '6':'2', 
            '7':'3', '8':'3', '9':'3', 
            '10':'4', '11':'4', '12':'4'}
territories = csv_dic('i/rebate_mdf_upload/auxiliary/territories.csv')
vendors = csv_dic('i/rebate_mdf_upload/auxiliary/vendors.csv')

############### Main ###############
t0 = time.clock()
with open(OUTPUT_CSV, 'w') as f:
    # Write Headers
    r = ','.join(['Division', 'Department', 'Territory', 'GL', 'Category', 
                  'Project', 'Vendor', 'Interco', 'Future', 'Month', 'Amount',
                  'GL Parent', 'Super Category', 'Region', 'Vendor Name', 'Quarter',
                  '\n'])
    f.write(r)
    
    for row, month, gl_split_header in product(rows, months, gl_split_headers):
        # Assign Month->Quarter
        quarter = quarters[month]
        # Calculate (Annual * Month * GL Split) Combination Total
        v = annual_forecast['2016_Annual'][row]
        monthly_v = monthly_seasonality[month][row]
        gl_split_v = gl_split[gl_split_header][row]
        new_monthly_v = v * monthly_v * gl_split_v
        
        # Extract GL Combination -> div, category, vendor, department
        gl = int(annual_forecast.iloc[row]['GL'])
        div, category, vendor, department = gl_split_header.split('-')
        if new_monthly_v != 0:
        
            # Assign GL Parent, Super Category, Vendor Name
            gl_parent = gl_parents.get(str(gl))
            super_category = categories.get(category)
            vendor_name = vendors.get(vendor)
        
            # Convert CAD->USD
            if div == '100':
                new_monthly_v = new_monthly_v * (1/usd_cad_fx)      # USD -> CAD

            # Rebate: Split Division -> Branches (Territory)
            if gl == 795110:
                if div == '100':
                    for div_100_branch_split_header in div_100_branch_split_headers:
                        branch_v = div_100_branch_split.iloc[row][div_100_branch_split_header]
                        new_monthly_branch_v = branch_v * new_monthly_v
                        if new_monthly_branch_v != 0:
                            # Assign Region
                            region = territories.get(div_100_branch_split_header)
                            r = ','.join([div, department, str(div_100_branch_split_header), 
                                          str(gl), category, '000000', vendor, '000', 
                                          '000000', month, str(new_monthly_branch_v), gl_parent,
                                          super_category, region, vendor_name, quarter, '\n'])
                            f.write(r)
                       
                elif div == '200':
                    for div_200_branch_split_header in div_200_branch_split_headers:
                        branch_v = div_200_branch_split.iloc[row][div_200_branch_split_header]
                        new_monthly_branch_v = branch_v * new_monthly_v
                        if new_monthly_branch_v != 0:
                            # Assign Region
                            region = territories.get(div_200_branch_split_header)
                            r = ','.join([div, department, str(div_200_branch_split_header), 
                                          str(gl), category, '000000', vendor, '000', 
                                          '000000', month, str(new_monthly_branch_v), gl_parent,
                                          super_category, region, vendor_name, quarter, '\n'])
                            f.write(r)
                       
            # All other GLs: Don't split Division -> Branches (Territory)
            else:
                region = 'Corporate'
                r = ','.join([div, department, '000', str(gl), category, 
                              '000000', vendor, '000', '000000', month, 
                              str(new_monthly_v), gl_parent, super_category, 
                              region, vendor_name, quarter, '\n'])
                f.write(r)

t1 = time.clock()
print 'rebate_mdf_upload_split.py complete.'
print t1-t0, 's'