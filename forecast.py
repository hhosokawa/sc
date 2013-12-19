import time
import numpy as np
import pandas as pd

""" io """

oracle_df = pd.read_csv('i\\forecast\\oracle_is.csv')
imputed_df = pd.read_csv('i\\forecast\\imputed_rev.csv')

output  = 'o\\Enrol Repo - %s.csv' % (time.strftime("%Y-%m-%d"))

############### Uitls ###############

############### forecast_main() ###############

def forecast_main():
    t0 = time.clock()
    oracle_df['Imputed Rev'] = 0

    print imputed_df.head(), '\n'
    print oracle_df.head()




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
