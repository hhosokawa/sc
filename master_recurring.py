import time
import numpy as np
import pandas as pd

output = 'o\\master_recurring.csv'
sales_df   = pd.read_csv('i\\master_recurring\\master_recurring.csv')
ref_df     = pd.read_csv('i\\master_recurring\\master_recurring_ref.csv')

renewal_dict = {}
yearly_renewal = {}

#################################################################################
## Function Definitions

def apply_renewal(series):
    return renewal_dict[series.name]

#################################################################################
## Main

def main():
    t0 = time.clock()

    # ref_df - rename headers
    ref_df.rename(columns={
        'Softchoice Master Master Number' : 'Softchoice Bill to Master Master Number',
        'Referral Fiscal Year' : 'Invoice date (SC FY)',
        'Referral Super Category' : 'Item Super Category',
        'Referral Revenue Total' : 'Total GP $'}, inplace = True)

    # sales_df = Obtain correct Headers
    filter_df = sales_df.ix[:, [
                'Softchoice Bill to Master Master Number',
                'Item Super Category',
                'Invoice date (SC FY)',
                'Total GP $']]

    # ref_df + sales_df
    combine_df = filter_df.append(ref_df)

    # Create Pivot Table
    pivot_df = pd.pivot_table(combine_df,
                            values = 'Total GP $',
                            rows = ['Item Super Category',
                                    'Softchoice Bill to Master Master Number'],
                            cols = ['Invoice date (SC FY)'],
                            aggfunc=np.sum)

    # Determine Renewal %
    for row in pivot_df.iterrows():
        numerator, denominator = 0., 0.
        super_master, dates = row
        for year in range(2005,2014):
            amt = dates[year]

            # Determine Demoninator
            if not np.isnan(amt):
                denominator += 1

                # Determine Numerator
                next_amt = dates[year+1]
                if next_amt > 0:
                    numerator += 1

        # Determine Renewal %
        try: renewal = numerator / denominator
        except ZeroDivisionError: renewal = 0
        renewal_dict[super_master] = renewal

    # Apply Renewal % to DF
    pivot_df['Renewal %'] = pivot_df.apply(apply_renewal, axis = 1)

    # Write to CSV
    pivot_df.to_csv(output)

    t1 = time.clock()
    return 'Process completed! Duration:', t1-t0

print main()
