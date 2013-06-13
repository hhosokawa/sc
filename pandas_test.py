import pandas as pd

datachunks = pd.read_csv(
             'C:/Portable Python 2.7/App/Scripts/i/database/2012-sales.csv',
             iterator=True, chunksize=1000)
data = pd.concat([chunk for chunk in datachunks], ignore_index=True)

def main():
    print data.head(), '\n'

if __name__ == '__main__':
    main()


"""
    data['category'] = data['revenue'].map(revenuecheck)

    grouped = data.groupby(['category', 'category 2'])
    pivot = grouped.sum()
    print pivot.head()
"""