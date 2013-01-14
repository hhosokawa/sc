import csv
import time
from fuzzywuzzy import fuzz
from collections import defaultdict
import numpy as np
import sys
from aux_reader import *

output = 'o\\14-Jan-13 - VSOE Old Analysis Sample.csv'
input0 = 'i\\VSOE\\Old_VSOE_Key.csv'
input1 = 'i\\VSOE\\VSOE 2012-2013 swbndl.csv'
input2 = 'i\\VSOE\\VSOE 2012-2013 swlic-swmtn.csv'
f = open("C:/Portable Python 2.7/App/Scripts/o/14-Jan-13 - VSOE Old sample.txt", "w")

#################################################################################
## Function Definition

def tree(): return defaultdict(tree)

# Compare Bundle Row to Lic/Mtn Row
def compare(bundlr, r):
    saname = (r['Product Description'] + ' ' +
              r['Product Title & Description'] + ' ' +
              r['Product Title'] + ' ' +
              r['Item Prodtype Description'] + ' ' +
              r['Item Prodtype Venprogram'] + ' ' +
              r['Item Category'] + ' ' +
              r['Product Publisher Name']).upper()
    inv_item = (r['Invoice Number'] + '-' + r['Item Number']  + '-' +
                r['Product Title & Description'] + '-' + r['Item Sell Price'])
    bundsell = float(bundlr['Item Sell Price'])
    sasell = float(r['Item Sell Price'])

    # Keyword Match
    if listinstring(bundlr['Keyword'], saname):
        if r['Item Revenue Recognition'] == 'SWLIC':
            bundlr['LIC Sample $'].append(sasell)
            bundlr['LIC Sample Inv'].append(inv_item)
        elif r['Item Revenue Recognition'] == 'SWMTN':
            bundlr['MTN Sample $'].append(sasell)
            bundlr['MTN Sample Inv'].append(inv_item)
        bundlr['Trigger'] = True
    return bundlr

# Compares KeywordList to Item Description String
def listinstring(list, string):
    for x in list:
        if x not in string: return False
    return True

# Writes Sample Bundle/Mtn/Lic Sample to Text
def totext(r):
    item = (r['Invoice Number'] + '-' + r['Item Number'] + '-' +
            r['Product Title & Description'] + '\n')
    f.write(item)
    for x in r['LIC Sample Inv']:
        f.write('LIC: ' + x + '\n')
    for y in r['MTN Sample Inv']:
        f.write('MTN: ' + y + '\n')
    f.write('\n')
    return

# Compares Median of Sample to other samples.
def med_analysis(r):
    r['Keyword'] = ' '.join(r['Keyword'])

    # Lic/ Mtn Population Comparison Apply %
    if len(r['LIC Sample $']) > 0:
        r['LIC Sample Pop'] = medianloop(r['LIC Sample $'])
        licamt = np.median(r['LIC Sample $'])
    else:
        r['LIC Sample Pop'] = 'No Samples'
    if len(r['MTN Sample $']) > 0:
        r['MTN Sample Pop'] = medianloop(r['MTN Sample $'])
        mtnamt = np.median(r['MTN Sample $'])
    else:
        r['MTN Sample Pop'] = 'No Samples'

    # Populate No Sample Median Sell prices
    if r['MTN Sample Pop'] == 'No Samples':
        mtnamt = float(r['Item Sell Price']) - np.median(r['LIC Sample $'])
    elif r['LIC Sample Pop'] == 'No Samples':
        licamt = float(r['Item Sell Price']) - np.median(r['MTN Sample $'])

    # Applies Percentage
    r['Percentage'] = mtnamt / (mtnamt+licamt)
    r['LIC Sample $'] = licamt
    r['MTN Sample $'] = mtnamt
    return r

# Compares x from Sample list to Median
def medianloop(samplelist):
    samplemedian = np.median(samplelist)
    lb = 0.85*samplemedian
    ub = 1.15*samplemedian
    count = 0.
    for x in samplelist:
        if lb <= x <= ub:
            count += 1.
    return float(count) / float(len(samplelist))

def acceptsample(r):
    if r['Trigger'] == False: return False
    else:
        if r['MTN Sample Pop'] == 'No Samples':
            if r['LIC Sample Pop'] <= 0.7: return False
            else: return True
        elif r['LIC Sample Pop'] == 'No Samples':
            if r['MTN Sample Pop'] <= 0.7: return False
            else: return True
        else:
            if r['LIC Sample Pop'] <= 0.7: return False 
            elif r['MTN Sample Pop'] <= 0.7: return False
            else: return True

def aftertxtclean(r):
    r['LIC Sample Inv'] = len(r['LIC Sample Inv'])
    r['MTN Sample Inv'] = len(r['MTN Sample Inv'])
    return r
#################################################################################
## dictionaries setup

bndldict = defaultdict(str)
sadict = tree()
#################################################################################
## Main Program

def main():
    t0 = time.clock()
    # Historic Keyword extraction
    hist_key = csv_dic(input0, 7)
    print 'Historic Key Extration Complete.', time.clock()-t0

    # Build SWBNDL Dict
    with open(input1) as i1:
        i1r = csv.DictReader(i1)
        for r in i1r:
            inv_item_key = r['Invoice Number'] + r['Item Number']
            if inv_item_key in hist_key:
                bndldict[r['Item Number']] =r
                bndldict[r['Item Number']]['Keyword'] = list(hist_key[inv_item_key])
                bndldict[r['Item Number']]['LIC Sample $'] = []
                bndldict[r['Item Number']]['LIC Sample Inv'] = []
                bndldict[r['Item Number']]['MTN Sample $'] = []
                bndldict[r['Item Number']]['MTN Sample Inv'] = []
                bndldict[r['Item Number']]['Percentage'] = 0.
                bndldict[r['Item Number']]['MTN Sample Pop'] = 0.
                bndldict[r['Item Number']]['LIC Sample Pop'] = 0.
                bndldict[r['Item Number']]['Trigger'] = False
                keys = tuple(bndldict[r['Item Number']].keys())
    print 'SWBNDL Dictionary Built', time.clock()-t0

    # Build STAND ALONE Dict
    with open(input2) as i2:
        i2r = csv.DictReader(i2)
        for r in i2r:
            ppn = r['Product Publisher Name']
            id = r['Invoice Number'] + '-' + r['Item Number']
            year = r['Invoice date (SC FY)']
            sadict[year][ppn][id] = r
    print 'STAND ALONE Dictionary Built', time.clock()-t0

    with open(output, 'wb') as o:
        writer = csv.writer(o)
        ow = csv.DictWriter(o, fieldnames=keys)
        writer.writerows([keys])

        # Scan SWMTN / SWLIC
        for ls in bndldict:
            bndlpublish = bndldict[ls]['Product Publisher Name']
            year = bndldict[ls]['Invoice date (SC FY)']
            try:
                for sa in sadict[year][bndlpublish]:
                    bundlsell = float(bndldict[ls]['Item Sell Price'])
                    sasell = float(sadict[year][bndlpublish][sa]['Item Sell Price'])
                    if ((sasell < bundlsell) and
                    (bndldict[ls]['Product Title & Description'][:10] in
                    sadict[year][bndlpublish][sa]['Product Title & Description'])):
                        bndldict[ls] = compare(bndldict[ls],
                                       sadict[year][bndlpublish][sa])
                for sa in sadict[str(int(year)-1)][bndlpublish]:
                    bundlsell = float(bndldict[ls]['Item Sell Price'])
                    sasell = float(sadict[str(int(year)-1)][bndlpublish][sa]['Item Sell Price'])
                    if ((sasell < bundlsell) and
                    (bndldict[ls]['Product Title & Description'][:10] in
                    sadict[str(int(year)-1)][bndlpublish][sa]['Product Title & Description'])):
                        bndldict[ls] = compare(bndldict[ls],
                                       sadict[str(int(year)-1)][bndlpublish][sa])
            except KeyboardInterrupt:
                print 'Aborting!'
                f.close()
                sys.exit()
            except:
                print 'No Samples Found for', bndlpublish, year
                continue

            if bndldict[ls]['Trigger'] == True:
                bndldict[ls] = med_analysis(bndldict[ls])

            # Output SWBNDL
            totext(bndldict[ls])
            ow.writerow(aftertxtclean(bndldict[ls]))
            print 'Bundle:', ls, '- Matched -', time.clock()
    t1 = time.clock()
    f.close()
    return 'Process completed! Duration:', t1-t0

print main()
