import csv
import time
import dateutil.parser as dparser
from dateutil.relativedelta import relativedelta
import collections
from datetime import datetime, timedelta
from decimal import Decimal

def csv_dic(filename, style=1):     # Converts CSV to dict
    reader = csv.reader(open(filename, "rb"))
    if style == 1: my_dict = dict((k, v) for k, v in reader)
    elif style == 6:
        my_dict = dict((k, (v1, v2, v3, v4, v5, v6))
                        for k, v1, v2, v3, v4, v5, v6 in reader)
    return my_dict

#################################################################################
## Update Item Inputs

output = 'o\\10-DEC-12 future billing.csv'
input1 = 'i\\SB - future billings dec.csv'  # Reminder: Manually do Renewal * 90%
input2 = 'i\\SB - contract repo - 10-dec-12.csv' # Reminder: delete Duplicates

enrollhistory = csv_dic('auxiliary\\MS Future Billing - Enroll History.csv', 6)
indirectpo = csv_dic('auxiliary\\MS PO ItemNumber Sell Price.csv') # Left 9

#################################################################################
## Class Declaration
def tree(): return collections.defaultdict(tree)

#################################################################################
## Function Definitions - Future Billings File

def header_add(header):             # Adds New Headers
    newfields = ['Custom Category A', 'Custom Category B', 'GP',
                 'Imputed Rev', 'Gross Rev', 'Custom Category C',
                 'Custom Category D', 'Div', 'Region',
                 'Scheduled Bill Month', 'Scheduled Bill Quarter']
    for newfield in newfields: header.add(newfield)
    return header

def manage(r):
    if r['Level'] == 'B':
        r['GP'] = Decimal(r['Extended Amount']) * Decimal(0.001458)
    elif r['Level'] == 'C':
        r['GP'] = Decimal(r['Extended Amount']) * Decimal(0.000525)
    elif r['Level'] == 'D':
        r['GP'] = Decimal(r['Extended Amount']) * Decimal(0.000417)
    else:
        r['GP'] = Decimal(r['Extended Amount']) * Decimal(0.001667)
    r['Gross Rev'] = r['GP']
    r['Imputed Rev'] = 0
    r['Custom Category C'] = 'EA Recurring'
    r['Custom Category D'] = 'Manage Fee'
    return r

def deploy(r):
    if r['Level'] == 'B':
        r['GP'] = Decimal(r['Extended Amount']) * Decimal(0.001458)
    elif r['Level'] == 'C':
        r['GP'] = Decimal(r['Extended Amount']) * Decimal(0.000525)
    elif r['Level'] == 'D':
        r['GP'] = Decimal(r['Extended Amount']) * Decimal(0.000417)
    else:
        r['GP'] = Decimal(r['Extended Amount']) * Decimal(0.001667)
    r['Gross Rev'] = r['GP']
    r['Imputed Rev'] = 0
    r['Custom Category C'] = 'EA Recurring'
    r['Custom Category D'] = 'Deploy Fee'
    return r

def helpdesk(r):
    if float(r['Extended Amount']) < 50000:
        r['GP'] = 25
    elif 50001 < float(r['Extended Amount']) < 400000:
        r['GP'] = 83.33
    elif 400001 < float(r['Extended Amount']) < 1000000:
        r['GP'] = 333.33
    elif 1000001 < float(r['Extended Amount']) < 6000000:
        r['GP'] = 1041.67
    elif 6000001 < float(r['Extended Amount']) < 20000000:
        r['GP'] = 2083.33
    elif float(r['Extended Amount']) > 20000001:
        r['GP'] = 3125
    r['Custom Category D'] = 'Help Desk Fee'
    r['Custom Category C'] = 'EA Recurring'
    r['Imputed Rev'] = 0
    r['Gross Rev'] = r['GP']
    return r

def gpcalc(r):                                # Calculates GP
    adc = 'Agreement Desktop Count'
    ir = 'Imputed Rev'
    if r[adc] == '': r[adc] = 0
    if r['Custom Category A'] == 'ESA 2.0':
        if (((int(r[adc]) < 750) and (r['Level'] == 'A'))
              or r['Level'] == 'A1'):
            return Decimal(0.03)
        elif (((int(r[adc]) > 750) and (r['Level'] == 'A'))
                or r['Level'] == 'A2'):
            return Decimal(0.0275)
        elif r['Level'] == 'B':
            return Decimal(0.025)
        else:
            return Decimal(0.0325)
    else:
        return Decimal(0.02)
    return

def notescalcESA2(r):                             # Calculates Item Notes
    adc = 'Agreement Desktop Count'
    if r[adc] == '': r[adc] = 0
    if r['Custom Category A'] == 'ESA 2.0':
        if (((int(r[adc]) < 750) and (r['Level'] == 'A'))
              or r['Level'] == 'A1'):
            return 'Level A1 SAM Service Req, 90%+ Activated'
        elif (((int(r[adc]) > 750) and (r['Level'] == 'A'))
                or r['Level'] == 'A2'):
            return 'Level A2 SAM Service Req, 90%+ Activated'
        elif r['Level'] == 'B':
            return 'Level B SAM Service Reqs, 90%+ Activated'
        elif r['Level'] == 'C':
            return 'Level C 90%+ Activated'
        elif r['Level'] == 'D':
            return 'Level D 90%+'
        else: return 'Level E 90%+'
    return

def refclean(r):                                # Clean MS Future Billings Data
    fb_enrol_set.add(r['Agreement Number'])
    esa2date = dparser.parse('10/30/2011')
    podate = dparser.parse(r['Purchase Order Date'])
    pocheck = (r['Ordered Under Purchase Order Number'] + ' '
               + r['Part Number'])
    r['Anniversary Year'] = r['Scheduled Bill Date'][-4:]

    if (pocheck in indirectpo):                 # Non-EA Indirect
        r['Custom Category A'] = 'NON-EA DIRECT'
        r['Custom Category B'] = r['Program']
        r['Custom Category C'] = ''
        r['Custom Category D'] = ''
        r['Imputed Rev'] = (Decimal(r['Quantity']) *  Decimal(indirectpo[pocheck]))
        r['GP'] = Decimal(r['Imputed Rev']) - Decimal(r['Extended Amount'])
        r['Gross Rev'] = r['Imputed Rev']
    else:
        if esa2date > podate:                   # ESA 2.0
            r['Custom Category A'] = 'ESA 2.0'
            agreementyear = (int(r['Scheduled Bill Date'][-4:]) -
                             int(r['Purchase Order Date'][-4:]) + 1)
            r['Custom Category B'] = 'Year ' + str(agreementyear)
            r['Custom Category C'] = notescalcESA2(r)
            r['Custom Category D'] = ''
            r['Imputed Rev'] = Decimal(r['Extended Amount'])
            r['GP'] = Decimal(r['Extended Amount']) * gpcalc(r)
            r['Gross Rev'] = r['GP']
        else:                                   # ESA 3.0 - MAJOR / CORPORATE
                                                # Custom Category B/C ID
            try:
                if (enrollmentcat[r['Agreement Number']] == 'Renew Contract' or
                r['Agreement Status'] == 'Renewal Assumption - Annual Bill * 90%'):
                    r['Custom Category B'] = 'Renew Contract'
                    r['Custom Category C'] = 'EA Renewal'
                else:
                    r['Custom Category B'] = 'New Contract'
                    r['Custom Category C'] = 'EA New'
            except:
                r['Custom Category B'] = 'Renew Contract'
                r['Custom Category C'] = 'EA Renewal'
            if (r['Primary Public Customer Number'] in majoraccts):
                r['Custom Category A'] = 'ESA 3.0 - MAJOR'
                r['Custom Category D'] = 'Contract Management'
                r['Imputed Rev'] = Decimal(r['Extended Amount'])
                r['GP'] = Decimal(r['Extended Amount']) * gpcalc(r)
                r['Gross Rev'] = r['GP']
            else:
                r['Custom Category A'] = 'ESA 3.0 - CORPORATE'
                r['Custom Category D'] = 'Transact Fee'
                r['Imputed Rev'] = Decimal(r['Extended Amount'])
                r['GP'] = Decimal(r['Extended Amount']) * gpcalc(r)
                r['Gross Rev'] = r['GP']
                                                # Absorb ESA 3.0 #'s
            esa3key = r['Agreement Number'] + ' - ' + r['Scheduled Bill Date']
            if esa3key in esa3dict:
                esa3dict[esa3key]['GP'] += r['GP']
                esa3dict[esa3key]['Imputed Rev'] += r['Imputed Rev']
                esa3dict[esa3key]['Gross Rev'] += r['Gross Rev']
            else: esa3dict[esa3key] = r
# HAVE TO READDRESS HOW DETERMINING DIVISION AND REGION
    try:
        r['Div'] = enrollmentdiv[r['Agreement Number']]
        r['Region'] = enrollmentregion[r['Agreement Number']]
    except:
        r['Div'] = 'N/A'
        r['Region'] = 'N/A'
    sbm = dparser.parse(r['Scheduled Bill Date']).date()
    r['Scheduled Bill Month'] = sbm.strftime("%Y-%m")
    sbqmonth = sbm.strftime("%m")
    r['Scheduled Bill Quarter'] = sbm.strftime("%Y-") + quarterperiod[sbqmonth]
    return r

#################################################################################
## Contract Repository Functions

# Cleans Contract Repository + Enrollment History -> Output Future Billing
# New Row, Old Row, Enrollment History Data, (Writer Objects)
def historydataparse(r, oldr, ehdata, o, writer, ow):
    dp = dparser.parse
    extrayrs = 0
    # Enrollment History - Date, Amount, True-Up, Add-on, Type
    eh_date, eh_amt, eh_trup, eh_addon, eh_refsource, eh_billtype = ehdata
    r['Div'] = oldr['Contract Division']
    r['Agreement Number'] = oldr['Contract Number']
    r['Program Offering Code'] = oldr['Contract Program Name']
    r['Agreement Start Date'] = dp(oldr['Contract Start Date']).date()
    r['Agreement Status'] = oldr['Contract Status']
    r['Agreement End Date'] = dp(oldr['Contract End Date']).date()
    r['Agreement Desktop Count'] = oldr['Contract Units']
    r['Scheduled Bill Date'] = dp(eh_date).date()
    r['Scheduled Bill Month'] = r['Scheduled Bill Date'].strftime("%Y-%m")
    sbqmonth = r['Scheduled Bill Date'].strftime("%m")
    r['Scheduled Bill Quarter'] = (r['Scheduled Bill Date'].strftime("%Y-") +
                                   quarterperiod[sbqmonth])
    r['Primary Customer Name'] = oldr['Master Name'].title()
    r['Primary Public Customer Number'] = oldr['Master Number']
    r['Level'] = oldr['Contract Level']
    r['Purchase Order Date'] = dp(oldr['Contract Create Date']).date()
    r['Imputed Rev'] = Decimal(eh_amt)
    r['Extended Amount'] = Decimal(eh_amt)
    try:
        if oldr['Master Divloc'] == '':
            r['Div'] = 'N/A'
            r['Region'] = enrollmentregion[r['Agreement Number']]
        else:
            r['Region'] = divregion[r['Div'] + oldr['Master Divloc']]
    except:
        r['Div'] = 'N/A'
        r['Region'] = 'N/A'

    if eh_refsource == 'ESA 2.0':           # ESA 2.0 Analysis
        r['Custom Category A'] = 'ESA 2.0'
        if eh_billtype == 'OAP3':
            r['Custom Category B'] = 'Year 3'
        elif eh_billtype == 'OAP2':
            r['Custom Category B'] = 'Year 2'
            extrayrs = 1
        else:
            yrdelta = (int(r['Scheduled Bill Date'].strftime("%Y")) -
                       int(r['Purchase Order Date'].strftime("%Y"))) + 1
            r['Custom Category B'] = 'Year ' + str(yrdelta)
            extrayrs = (int(r['Agreement End Date'].strftime("%Y")) -
                        int(r['Scheduled Bill Date'].strftime("%Y"))) - 1
        r['Custom Category C'] = notescalcESA2(r)
        r['GP'] = r['Imputed Rev'] * gpcalc(r)
        r['Gross Rev'] = r['GP']
        ow.writerow(r)
    elif eh_refsource == 'ESA 3.0':         # ESA 3.0 CORP/MAJOR Analysis
        try:
            if enrollmentcat[r['Agreement Number']] == 'Renew Contract':
                r['Custom Category B'] = 'Renew Contract'
                r['Custom Category C'] = 'EA Renewal'
            else:
                r['Custom Category B'] = 'New Contract'
                r['Custom Category C'] = 'EA New'
        except:
            r['Custom Category B'] = 'Renew Contract'
            r['Custom Category C'] = 'EA Renewal'

        if (r['Primary Public Customer Number'] in majoraccts):
            r['Custom Category A'] = 'ESA 3.0 - MAJOR'
            r['Custom Category D'] = 'Contract Management'
            r['GP'] = Decimal(r['Extended Amount']) * gpcalc(r)
        else:
            r['Custom Category A'] = 'ESA 3.0 - CORPORATE'
            r['Custom Category D'] = 'Transact Fee'
            r['GP'] = Decimal(r['Extended Amount']) * gpcalc(r)
        r['Gross Rev'] = r['GP']
        ow.writerow(r)
        calcfees(r, o, writer, ow)
        extrayrs = (int(r['Agreement End Date'].strftime("%Y")) -
                    int(r['Scheduled Bill Date'].strftime("%Y"))) - 1
                                        # Calculate Timing of Enrollment Cycle
    if extrayrs == 0:                   # Enrollment up for Renewal
        addrenewal(r, o, writer, ow)
    else:                               # Enrollment has X years remaining (Max 2)
        for yr in range(extrayrs):
            if yr > 1: break
            if r['Custom Category A'] == 'ESA 2.0':
                newcatb = (r['Custom Category B'][:4] + ' ' +
                           str(int(r['Custom Category B'][-1:]) + 1))
                r['Custom Category B'] = newcatb
            r['Scheduled Bill Date'] = r['Scheduled Bill Date'] + relativedelta(years=+1)
            r['Scheduled Bill Month'] = r['Scheduled Bill Date'].strftime("%Y-%m")
            sbqmonth = r['Scheduled Bill Date'].strftime("%m")
            r['Scheduled Bill Quarter'] = (r['Scheduled Bill Date'].strftime("%Y-") +
                                           quarterperiod[sbqmonth])
            r['Ordered Under Purchase Order Number'] = yr
            ow.writerow(r)
            calcfees(r, o, writer, ow)
            if r['Scheduled Bill Date']+relativedelta(years=+1) > r['Agreement End Date']:
                addrenewal(r, o, writer, ow)
    return

def addrenewal(r, o, writer, ow):
    r['Agreement Start Date'] = r['Agreement End Date']
    r['Purchase Order Date'] = r['Agreement End Date']
    r['Scheduled Bill Date'] = r['Agreement End Date']
    r['Scheduled Bill Month'] = r['Scheduled Bill Date'].strftime("%Y-%m")
    sbqmonth = r['Scheduled Bill Date'].strftime("%m")
    r['Scheduled Bill Quarter'] = (r['Scheduled Bill Date'].strftime("%Y-") +
                                   quarterperiod[sbqmonth])
    r['Agreement End Date'] = r['Agreement End Date'] + relativedelta(years=+1)
    r['Imputed Rev'] = r['Imputed Rev'] * Decimal(0.9)
    r['Extended Amount'] = r['Imputed Rev']
    r['Custom Category B'] = 'Renew Contract'
    r['Custom Category C'] = 'EA Renewal'
    r['Agreement Status'] = 'Renewal Assumption - Annual Bill * 90%'
    if (r['Primary Public Customer Number'] in majoraccts):
        r['Custom Category A'] = 'ESA 3.0 - MAJOR'
        r['Custom Category D'] = 'Contract Management'
        r['GP'] = r['Extended Amount'] * gpcalc(r)
        r['Gross Rev'] = r['GP']
        ow.writerow(r)
        calcfees(r, o, writer, ow)
    else:
        r['Custom Category A'] = 'ESA 3.0 - CORPORATE'
        r['Custom Category D'] = 'Transact Fee'
        r['GP'] = r['Extended Amount'] * gpcalc(r)
        r['Gross Rev'] = r['GP']
        ow.writerow(r)
        calcfees(r, o, writer, ow)
        r['Custom Category D'] = 'On-Time Renewal Fee'
        r['Imputed Rev'] = 0
        r['GP'] = r['Extended Amount'] * ontimerenewal(r)
        r['Gross Rev'] = r['GP']
        ow.writerow(r)
    return

def calcfees(r, o, writer, ow):
    originalcat = r['Custom Category D']
    originalcatc = r['Custom Category C']
    originalsbd = r['Scheduled Bill Date']
    originalsbm = r['Scheduled Bill Month']
    originalsbq = r['Scheduled Bill Quarter']
    for x in range(12):
        if r['Custom Category A'] == 'ESA 3.0 - CORPORATE':
            ow.writerow(manage(r))
            ow.writerow(deploy(r))
        elif r['Custom Category A'] == 'ESA 3.0 - MAJOR':
            ow.writerow(helpdesk(r))
        r['Scheduled Bill Date'] = r['Scheduled Bill Date'] + relativedelta(months=+1)
        r['Scheduled Bill Month'] = r['Scheduled Bill Date'].strftime("%Y-%m")
        sbqmonth = r['Scheduled Bill Date'].strftime("%m")
        r['Scheduled Bill Quarter'] = (r['Scheduled Bill Date'].strftime("%Y-") +
                                       quarterperiod[sbqmonth])
    r['Imputed Rev'] = r['Extended Amount']
    r['Custom Category D'] = originalcat
    r['Custom Category C'] = originalcatc
    r['Scheduled Bill Date'] = originalsbd
    r['Scheduled Bill Month'] = originalsbm
    r['Scheduled Bill Quarter'] = originalsbq
    return

def ontimerenewal(r):
    adc = 'Agreement Desktop Count'
    if r[adc] < 750: return Decimal(0.02)
    elif 750 < r[adc] < 2399: return Decimal(0.075)
    elif 2400 < r[adc] < 5999: return Decimal(0.005)
    elif 6000 < r[adc] < 14999: return Decimal(0.0025)
    else: return Decimal(0.0025)

#################################################################################
## Dictionary Pictionary Jars

majoraccts = csv_dic('auxiliary\\enrol - major customers.csv')
enrollmentcat = csv_dic('auxiliary\\MS Enrollment - Category.csv')
enrollmentdiv = csv_dic('auxiliary\\Ms Enrollment - Div.csv')
enrollmentregion = csv_dic('auxiliary\\Ms Enrollment - Region.csv')
divregion = csv_dic('auxiliary\\div-region.csv')
esa3dict = tree()
enroltree = tree()
fb_enrol_set = set()
quarterperiod = {'01': 'Q1', '02': 'Q1', '03': 'Q1', '04': 'Q2', '05': 'Q2',
                 '06': 'Q2', '07': 'Q3', '08': 'Q3', '09': 'Q3', '10': 'Q4',
                 '11': 'Q4', '12': 'Q4'}

#################################################################################
## Main

def main():
    t0 = time.clock()
    header = set()                  # Get all input headers -> output header
    with open(input1) as i1: header.update(csv.DictReader(i1).fieldnames)
    header = header_add(header)
    header = tuple(header)

# Analyze Input 1: Explore.ms Future Billings Report -> Output
    with open(output, 'wb') as o:
        writer = csv.writer(o)
        writer.writerows([header])
        with open(input1) as i1:
            i1r = csv.DictReader(i1)
            ow = csv.DictWriter(o, fieldnames=header)
            for r in i1r:
                ow.writerow(refclean(r))
                                    # On-Time Renewal
                if (r['Custom Category A'] == 'ESA 3.0 - CORPORATE' and
                    r['Custom Category B'] == 'Renew Contract' and
                    r['Agreement Status'] == 'Renewal Assumption - Annual Bill * 90%'):
                    adc = 'Agreement Desktop Count'
                    if (r['Level'] == 'A') and (int(r[adc]) < 750):
                        r['GP'] = Decimal(0.085) * Decimal(r['Extended Amount'])
                    elif (r['Level'] == 'A') and (int(r[adc]) > 750):
                        r['GP'] = Decimal(0.0425) * Decimal(r['Extended Amount'])
                    elif r['Level'] == 'B': r['GP'] = (
                        Decimal(0.035) * Decimal(r['Extended Amount']))
                    elif r['Level'] == 'C': r['GP'] = (
                        Decimal(0.0175) * Decimal(r['Extended Amount']))
                    else: r['GP'] = (
                        Decimal(0.0125) * Decimal(r['Extended Amount']))
                    r['Gross Rev'] = r['GP']
                    r['Custom Category D'] = 'On-Time Renewal Fee'
                    r['Imputed Rev'] = 0
                    ow.writerow(r)

                                    # Inserts Monthly Billings for ESA 3.0
            for month_esa3 in esa3dict:
                if esa3dict[month_esa3]['Imputed Rev'] == 0:
                    pass
                else: esa3dict[month_esa3]['Extended Amount'] = (
                    esa3dict[month_esa3]['Imputed Rev'])
                covstart = dparser.parse(esa3dict[month_esa3]['Coverage Start Date'])
                covend = dparser.parse(esa3dict[month_esa3]['Coverage End Date'])
                billdate = dparser.parse(esa3dict[month_esa3]['Scheduled Bill Date'])
                deltadays = covend - covstart

                if deltadays > timedelta(days=360):
                    for x in range(12):
                        if esa3dict[month_esa3]['Custom Category A'] == (
                            'ESA 3.0 - CORPORATE'):
                            ow.writerow(manage(esa3dict[month_esa3]))
                            ow.writerow(deploy(esa3dict[month_esa3]))
                        else:
                            ow.writerow(helpdesk(esa3dict[month_esa3]))
                        billdate = billdate + relativedelta(months=+1)
                        esa3dict[month_esa3]['Scheduled Bill Date'] = billdate.date()
                        sbm = esa3dict[month_esa3]['Scheduled Bill Date'].strftime("%Y-%m")
                        esa3dict[month_esa3]['Scheduled Bill Month'] = sbm
                        sbqmonth = esa3dict[month_esa3]['Scheduled Bill Date'].strftime("%m")
                        r['Scheduled Bill Quarter'] = (billdate.strftime("%Y-") +
                                                       quarterperiod[sbqmonth])

# Analyze Input 2: Contract Repository -> Output
        with open(input2) as i2:
            dp = dparser.parse
            i2r = csv.DictReader(i2)
            for oldr in i2r:
                if (oldr['Contract Number'] in enrollhistory and
                    oldr['Contract Number'] not in fb_enrol_set):
                    r = dict.fromkeys(header)
                    ehdata = enrollhistory[oldr['Contract Number']]
                    historydataparse(r, oldr, ehdata, o, writer, ow)
        t1 = time.clock()
        return 'Process completed! Duration:', t1-t0

print main()

