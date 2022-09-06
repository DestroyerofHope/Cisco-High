# -*- coding: utf-8 -*-
"""
Created on Thu Aug 11 21:53:22 2022

@author: jasoli3
"""

def classifyAddress(address):
    statelist = {'Alabama':'AL','Alaska':'AK','Arizona':'AZ','Arkansas':'AR','California':'CA','Colorado':'CO','Connecticut':'CT','Delaware':'DE','Florida':'FL','Georgia':'GA','Hawaii':'HI','Idaho':'ID','Illinois':'IL','Indiana':'IN',
    'Iowa':'IA','Kansas':'KS','Kentucky':'KY','Louisiana':'LA','Maine':'ME','Maryland':'MD','Massachusetts':'MA','Michigan':'MI','Minnesota':'MN','Mississippi':'MS','Missouri':'MO','Montana':'MT','Nebraska':'NE','Nevada':'NV',
    'New Hampshire':'NH','New Jersey':'NJ','New Mexico':'NM','New York':'NY','North Carolina':'NC','North Dakota':'ND','Ohio':'OH','Oklahoma':'OK','Oregon':'OR','Pennsylvania':'PA','Rhode Island':'RI','South Carolina':'SC',
    'South Dakota':'SD','Tennessee':'TN','Texas':'TX','Utah':'UT','Vermont':'VT','Virginia':'VA','Washington':'WA','West Virginia':'WV','Wisconsin':'WI','Wyoming':'WY','Alberta':'AB','British Columbia':'BC','Manitoba':'MB',
    'New Brunswick':'NB','Newfoundland and Labrador':'NL','Nova Scotia':'NS','Ontario':'ON','Prince Edward Island':'PE','Quebec':'QC','Saskatchewan':'SK'}
    state, city, zipcode, country = '', '', '',''
    actualaddress = ''
    address = address.split(' ')
    for i in range(len(address)):
        address[i] = address[i].replace(',', '')
        address[i] = address[i].replace('.', '')
    for name in statelist:
        for j in range(len(address)):
            if name.lower() == address[j].lower() or statelist[name].lower() == address[j].lower():
                state = name
                statenum = j
    if state != '':
        city = address[statenum-1]
        for i in range(statenum, len(address)):
            if address[i].lower() in ['us', 'usa']:
                country = 'United States'
            if address[i].lower() == 'united' and address[i+1].lower() == 'states':
                country = 'United States'
            if address[i].lower() in ['canada', 'ca']:
                country = 'Canada'
            try:
                if len(str(int(address[i]))) == 5:
                    zipcode = str(address[i])
            except:
                None
            letternum, intnum = 0,0
            if len(address[i]) == 6:
                letters = [*address[i]]
                for item in letters:
                    try:
                        if type(int(item)) == int:
                            intnum += 1
                    except:
                        letternum += 1
                if letternum == 3 and intnum == 3:
                    zipcode = address[i]
            if len(address[i]) == 3 and len(address[i+1]) == 3:
                letters = [*(address[i]+address[i+1])]
                for item in letters:
                    try:
                        if type(int(item)) == int:
                            intnum += 1
                    except:
                        letternum += 1
                if letternum == 3 and intnum == 3:
                    zipcode = address[i] + address[i+1]
        address = address[0:statenum-1]
        for item in address:
            actualaddress += item + ' '
    else:
        for item in address:
            actualaddress += item + ' '
            
    return actualaddress, city, state, zipcode, country

print(classifyAddress('1366 Reale Avenue, Saint Louis, Missouri, K2T 0L9, United States'))