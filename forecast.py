import time
import pandas as pd
from aux_reader import *

""" io """

oracle_df = pd.read_csv('i\\2013oracle.csv')

output  = 'o\\Enrol Repo - %s.csv' % (time.strftime("%Y-%m-%d"))

############### Uitls ###############

############### enrol_repo_main() ###############

def forecast_main():
    t0 = time.clock()
    print oracle_df
    t1 = time.clock()
    print 'forecast_main() completed. Duration:', t1-t0
    return

if __name__ == '__main__':
    forecast_main()

