import time
import numpy as np
import pandas as pd

""" io """
oracle_df =     pd.read_csv('i\\forecast\\oracle_is.csv')
imputed_df =    pd.read_csv('i\\forecast\\imputed_rev.csv')

output  =       'o\\forecast - %s.csv' % (time.strftime("%Y-%m-%d"))

""" pictionary """
oracle_bi_dic = {'Canada' : '100 - Canada',
    'United States' : '200 - US',
    'N/A' : 'Default',
    'Other' : 'Default',
    'Undefined' : 'Default',
    'Managed Services' : 'Services',
    'Professional Services (INACTIVE)' : 'Services'}

bi_oracle_header = {'Division' : 'Div',
    'Calendar Year' : 'Year',
    'Fiscal Period' : 'Period',
    'Fiscal Quarter' : 'Quarter',
    'Super Category @ Order Date' : 'Super Cat',
    'Imputed Revenue (includes Freight)' : 'Imputed Rev'}

############### oracle ###############

def field_margin(series):
    cat_C = series['Cat C']
    if ('Vendor Rebate Revenue' in cat_C
    or 'Co-op Marketing Revenue' in cat_C):
        return 0
    else:
        return series['Amount']

def net_rev(series):
    cat_A = series['Cat A']
    if 'Net sales' in cat_A:
        return series['Amount']
    else:
        return 0

def rebate(series):
    cat_C = series['Cat C']
    if 'Vendor Rebate Revenue' in cat_C:
        return series['Amount']
    else:
        return 0

def mdf(series):
    cat_C = series['Cat C']
    if 'Co-op Marketing Revenue' in cat_C:
        return series['Amount']
    else:
        return 0

def oracle_clean():
    oracle_df['Imputed Rev'] = 0
    oracle_df['MDF Rev'] = oracle_df.apply(mdf, axis = 1)
    oracle_df['FM'] = oracle_df.apply(field_margin, axis = 1)
    oracle_df['Net Rev'] = oracle_df.apply(net_rev, axis = 1)
    oracle_df['Rebate Rev'] = oracle_df.apply(rebate, axis = 1)
    print 'oracle_clean() completed.'

############### bi ###############

def add_q(series):
    qtr = str(series['Fiscal Quarter'])
    qtr = 'Q' + qtr
    return qtr

def cat_replace(series):
    super_cat = str(series['Super Category @ Order Date'])
    if super_cat in oracle_bi_dic:
       series['Super Category @ Order Date'] = oracle_bi_dic[super_cat]
    return series['Super Category @ Order Date']

def div_replace(series):
    div = str(series['Division'])
    series['Division'] = oracle_bi_dic[div]
    return series['Division']

def add_ms_agency(series):
    super_cat = series['Super Category @ Order Date']
    sale_ref = series['Sale or Referral']
    if super_cat == 'Microsoft' and sale_ref == 'Referral':
        series['Super Category @ Order Date'] = 'Microsoft Agency Fees'
    return series['Super Category @ Order Date']

def bi_clean():
    global imputed_df
    imputed_df['FM'] = 0
    imputed_df['Amount'] = 0
    imputed_df['Net Rev'] = 0
    imputed_df['MDF Rev'] = 0
    imputed_df['Rebate Rev'] = 0
    imputed_df['Fiscal Quarter'] = imputed_df.apply(add_q, axis = 1)
    imputed_df['Super Category @ Order Date'] = imputed_df.apply(
                                                cat_replace, axis = 1)
    imputed_df['Division'] = imputed_df.apply(div_replace, axis = 1)
    imputed_df['Super Category @ Order Date'] = imputed_df.apply(
                                                add_ms_agency, axis = 1)
    imputed_df = imputed_df.rename(columns = bi_oracle_header)
    print 'bi_clean() completed.'

def bi_duplicate_consol():
    bi_consol_df = imputed_df.copy()
    bi_consol_df['Div'] = '* - Consolidated'
    print 'bi_duplicate_consol() completed.'
    return bi_consol_df

############### forecast_main() ###############

if __name__ == '__main__':
    t0 = time.clock()
    oracle_clean()
    bi_clean()
    bi_consol_df = bi_duplicate_consol()

    concat_df = pd.concat([oracle_df, imputed_df, bi_consol_df])
    concat_df.to_csv(output, index = False)
    t1 = time.clock()
    print 'forecast_main() completed. Duration:', t1-t0