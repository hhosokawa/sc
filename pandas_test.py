import pandas as pd
import numpy as np

data = pd.read_csv('c:/users/hhos/Downloads/fortune1000.csv')
famous = {'3Com': 'Cool'}

def revenuecheck(revenue):
    if 0 < revenue < 1000:
        return 'A'
    elif 1001 < revenue < 2000:
        return 'B'
    elif 2001 < revenue < 3000:
        return 'C'
    else:
        return 'Tosin Abasi'

def revenuecheck2(rev):
    if 0 < rev < 2000:
        return 'Super 1'
    elif 2001 < rev < 3000:
        return 'Super 2'
    elif 3001 < rev < 4000:
        return 'Mega 3'
    else:
        return 'Mega Tosin Abasi'

def main():
    print data.head(), '\n'
    data['category'] = data['revenue'].map(revenuecheck)
    data['category 2'] = data['revenue'].map(revenuecheck2)

    print 'Edit 1'
    print data.head(), '\n'

    grouped = data.groupby(['category', 'category 2'])
    pivot = grouped.sum()
    print pivot.head()




if __name__ == '__main__':
    main()
