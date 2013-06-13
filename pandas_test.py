import pandas as pd
import numpy as np

# Input
data = pd.read_csv('C:/sc/i/database/2012-sales.csv')
output = 'C:/sc/o/outputtest.csv'

# Select Columns
headers =   ['Bill to Master NAICS Description',
            'Softchoice DivLoc Des',
            'Invoice date (SC FM)',
            'Total Sale price']

# Pivot: Val, Rows, Cols
pivot_val  =    'Total Sale price'
pivot_rows =    ['Bill to Master NAICS Description',
                'Softchoice DivLoc Des']
pivot_cols =    ['Invoice date (SC FM)']

def main():
    # Select Columns
    filterdata = data.ix[:, headers]

    # Select Pivot Rows / Columns
    pivotdata = pd.pivot_table(filterdata,
                        values = pivot_val,
                        rows = pivot_rows,
                        cols = pivot_cols,
                        aggfunc = np.sum)

    # Output to CSV
    pivotdata.to_csv(output)
    return

if __name__ == '__main__':
    main()


"""
    grouped = filterdata.groupby([
        'Bill to Master NAICS Description',
        'Softchoice DivLoc Des'])

    print grouped.sum()[:10]
"""
