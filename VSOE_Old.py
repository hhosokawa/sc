import csv
import time
from fuzzywuzzy import fuzz
from collections import defaultdict
import numpy as np
import sys

output = 'o\\10-Jan-13 - VSOE Old Analysis.csv'
input1 = 'i\\2011-2012 SWBNDL.csv'
input2 = 'i\\2010-2012 SWLIC-SWMTN.csv'
f = open("C:/Portable Python 2.7/App/Scripts/o/10-Jan-13 - VSOE Old.txt", "w")

#################################################################################
## Function Definition

def tree(): return defaultdict(tree)

# Extract Keywords from Bundle
def keyword(r):
    keywordlist = []
    name = (r['Product Description'] + ' ' +
            r['Product Title & Description'] + ' ' +
            r['Product Title'] + ' ' +
            r['Item Prodtype Description'] + ' ' +
            r['Item Prodtype Venprogram']).upper()
    if r['Product Publisher Name'] == 'MICROSOFT':
        if 'SVR' in name: keywordlist.append('SVR')
        elif 'APP' in name: keywordlist.append('APP')
        elif 'SYS' in name:  keywordlist.append('SYS')
        elif 'SLG' in name:  keywordlist.append('SLG')
        if '3YR' in name: keywordlist.append('3YR')
        elif '2YR' in name: keywordlist.append('2YR')
        elif '1YR' in name: keywordlist.append('1YR')
        elif 'MONTH' in name: keywordlist.append('MONTH')
        elif '3' in name: keywordlist.append('3')
        elif '2' in name: keywordlist.append('2')
        elif '1' in name: keywordlist.append('1')
        if 'MVL C' in name: keywordlist.append('MVL C')
        elif 'MVL B' in name: keywordlist.append('MVL B')
        elif 'MVL A' in name: keywordlist.append('MVL A')
        elif 'MVL EDU A' in name: keywordlist.append('MVL EDU A')
        elif 'MVL D' in name: keywordlist.append('MVL D')
        elif 'MS SELECT PLUS A' in name: keywordlist.append('MS SELECT PULS A')
        elif 'MS SELECT PLUS C' in name: keywordlist.append('MS SELECT PULS C')
        elif 'MOL GOV' in name: keywordlist.append('MOL GOV')
        elif 'MOL CHAR' in name: keywordlist.append('MOL CHAR')
        elif 'VAL' in name: keywordlist.append('VAL')
        elif 'MVL CSLP D' in name: keywordlist.append('MVL CSLP D')
        elif 'MOL EDU' in name: keywordlist.append('MOL EDU')
        elif 'MOL BUS' in name: keywordlist.append('MOL BUS')
        elif 'MVL RBH D' in name: keywordlist.append('MVL RBH D')
        elif 'ADD' in name: keywordlist.append('ADD')
        elif 'MVL DISO' in name: keywordlist.append('MVL DISO')
        elif 'MOL VAL' in name: keywordlist.append('MOL VAL')
        elif 'MS SELECT PLUS D' in name: keywordlist.append('MS SELECT PLUS D')
        elif 'MS SELECT PLUS B' in name: keywordlist.append('MS SELECT PLUS B')
        elif 'VALUE' in name: keywordlist.append('VALUE')
        if 'ANN PAY' in name: keywordlist.append('ANN PAY')
        elif 'REMAINING' in name: keywordlist.append('REMAINING')
        elif 'DEVICE' in name: keywordlist.append('DEVICE')
        elif 'USER' in name: keywordlist.append('USER')
        elif 'UPGRADE' in name: keywordlist.append('UPGRADE')
        elif 'GIA' in name: keywordlist.append('GIA')
    elif r['Product Publisher Name'] == 'SYMANTEC':
        if 'BASIC' in name: keywordlist.append('BASIC')
        elif 'ESSENTIAL' in name: keywordlist.append('ESSENTIAL')
        if 'REWARD' in name: keywordlist.append('REWARD')
        elif 'GOVERNMENT' in name: keywordlist.append('GOVERNMENT')
        elif 'DISO' in name: keywordlist.append('DISO')
        elif 'EXPRESS' in name: keywordlist.append('EXPRESS')
        elif 'ACADEMIC' in name: keywordlist.append('ACADEMIC')
        if 'LEVEL A' in name: keywordlist.append('LEVEL A')
        elif 'LEVEL B' in name: keywordlist.append('LEVEL B')
        elif 'LEVEL C' in name: keywordlist.append('LEVEL C')
        elif 'LEVEL D' in name: keywordlist.append('LEVEL D')
        elif 'LEVEL E' in name: keywordlist.append('LEVEL E')
        elif 'LEVEL F' in name: keywordlist.append('LEVEL F')
        elif 'LEVEL S' in name: keywordlist.append('LEVEL S')
    elif r['Product Publisher Name'] == 'QUEST':
        if 'DISO' in name: keywordlist.append('DISO')
        elif 'NORMAL' in name: keywordlist.append('NORMAL')
    elif r['Product Publisher Name'] == 'SOPHOS':
        if 'GOVERNMENT' in name: keywordlist.append('GOVERNMENT')
        elif 'NORMAL' in name: keywordlist.append('NORMAL')
        elif 'EDUCATION' in name: keywordlist.append('EDUCATION')
        if '1 YEAR' in name: keywordlist.append('1')
    elif r['Product Publisher Name'] == 'NOVELL':
        if 'MLA' in name: keywordlist.append('MLA')
    return keywordlist

# Compare Bundle Row to Lic/Mtn Row
def compare(bundlr, r):
    saname = (r['Product Description'] + ' ' +
              r['Product Title & Description'] + ' ' +
              r['Product Title'] + ' ' +
              r['Item Prodtype Description'] + ' ' +
              r['Item Prodtype Venprogram']).upper()
    lsproddesc = bundlr['Product Title'].upper()
    saproddesc = r['Product Title'].upper()
    inv_item = (r['Invoice Number'] + '-' + r['Item Number']  + '-' +
                r['Product Title & Description'] + '-' + r['Item Sell Price'])
    bundsell = float(bundlr['Item Sell Price'])
    sasell = float(r['Item Sell Price'])

    # Keyword + Product Match 95
    if ((listinstring(bundlr['Keyword'], saname) and
        fuzz.token_set_ratio(lsproddesc, saproddesc) >= 90 and
        len(bundlr['Keyword']) > 0 and
        (0.1*bundsell <= sasell <= bundsell)) or

        # OR Product Match 100
        (fuzz.token_set_ratio(lsproddesc, saproddesc) >= 95 and
        (0.1*bundsell <= sasell <= bundsell))):
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
            if r['LIC Sample Pop'] < 0.8: return False
            else: return True
        elif r['LIC Sample Pop'] == 'No Samples':
            if r['MTN Sample Pop'] < 0.8: return False
            else: return True
        else:
            if (r['LIC Sample Pop'] or r['MTN Sample Pop']) < 0.8: return False
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

    # Build SWBNDL Dict
    with open(input1) as i1:
        i1r = csv.DictReader(i1)
        for r in i1r:
            bndldict[r['Item Number']] =r
            bndldict[r['Item Number']]['Keyword'] = keyword(r)
            bndldict[r['Item Number']]['LIC Sample $'] = []
            bndldict[r['Item Number']]['LIC Sample Inv'] = []
            bndldict[r['Item Number']]['MTN Sample $'] = []
            bndldict[r['Item Number']]['MTN Sample Inv'] = []
            bndldict[r['Item Number']]['Percentage'] = 0.
            bndldict[r['Item Number']]['MTN Sample Pop'] = 0.
            bndldict[r['Item Number']]['LIC Sample Pop'] = 0.
            bndldict[r['Item Number']]['Trigger'] = False
            keys = tuple(bndldict[r['Item Number']].keys())
    print 'SWBNDL Dictionary Built', time.clock()

    # Build STAND ALONE Dict
    with open(input2) as i2:
        i2r = csv.DictReader(i2)
        for r in i2r:
            ppn = r['Product Publisher Name']
            id = r['Invoice Number'] + '-' + r['Item Number']
            year = r['Invoice date (SC FY)']
            sadict[year][ppn][id] = r

    print 'STAND ALONE Dictionary Built', time.clock()

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
                    if ((0.05*bundlsell <= sasell < bundlsell) and
                    (bndldict[ls]['Product Title & Description'][:20] in
                    sadict[year][bndlpublish][sa]['Product Title & Description'])):
                        bndldict[ls] = compare(bndldict[ls],
                                       sadict[year][bndlpublish][sa])
                for sa in sadict[str(int(year)-1)][bndlpublish]:
                    bundlsell = float(bndldict[ls]['Item Sell Price'])
                    sasell = float(sadict[str(int(year)-1)][bndlpublish][sa]['Item Sell Price'])
                    if ((0.05*bundlsell <= sasell < bundlsell) and
                    (bndldict[ls]['Product Title & Description'][:20] in
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
            if acceptsample(bndldict[ls]):
                totext(bndldict[ls])
                ow.writerow(aftertxtclean(bndldict[ls]))
                print 'Bundle:', ls, '- Matched -', time.clock()
            else:
                print 'Bundle:', ls, '- No Matches -', time.clock()
    t1 = time.clock()
    f.close()
    return 'Process completed! Duration:', t1-t0

print main()
