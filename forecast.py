import time
import numpy as np
import pandas as pd

""" io """
oracle_df =     pd.read_csv('i\\forecast\\oracle_is.csv')
imputed_df =    pd.read_csv('i\\forecast\\imputed_rev.csv')

output  =       'o\\Enrol Repo - %s.csv' % (time.strftime("%Y-%m-%d"))

""" pictionary """
oracle_bi_dic = {}


############### utils ###############

def net_rev(series):
    cat_A = series['Cat A']
    if 'Net sales' in cat_A:
        return series['Amount']
    else:
        return 0

def oracle_clean():
    oracle_df['Imputed Rev'] = 0
    oracle_df['Net Rev'] = oracle_df.apply(net_rev, axis = 1)


def add_q(series):
    qtr = str(series['Fiscal Quarter'])
    qtr = 'Q' + qtr
    return qtr

def bi_clean():
    imputed_df['Amount'] = 0
    imputed_df['Net Rev'] = 0
    imputed_df['Fiscal Quarter'] = imputed_df.apply(add_q, axis = 1)
    print imputed_df.head()
    raw_input('...')


############### forecast_main() ###############

def forecast_main():
    t0 = time.clock()
    oracle_clean()
    bi_clean()







    t1 = time.clock()
    print 'forecast_main() completed. Duration:', t1-t0
    return

if __name__ == '__main__':
    forecast_main()




"""
    oracle_pt = pd.pivot_table(oracle_df,
                               values = 'Amount',
                               rows = ['Div', 'Super Cat'],
                               cols=['Year'],
                               aggfunc=np.sum)
"""


