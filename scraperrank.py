# coding=utf-8

import sys
import re
import warnings
from robobrowser import RoboBrowser

if not sys.warnoptions:
    warnings.simplefilter("ignore")


def num_there(s):
    return any(i.isdigit() for i in s)


# List of Subjects
subj = ['1', '2', '3', '4', '5', '6', '7', '8']
ch = 0

# Input for Branch and USNs

college = ["1MV", "1CR", "1RN", "1VA", "1VE", "1BY"]
year = input('Enter the year\n')
branch = input('Please enter the branch\n').upper()
low = int(input('Enter starting USN\n'))
if low >= 400:
    dip = 'Y'
else:
    dip = 'N'
# increment last USN to aid looping
high = int(input('Enter last USN\n')) + 1
semc = input('Enter the Semester\n')

subcode = 52
iloop = 8

if semc == '1':
    cycle = input('Enter the Cycle\n').upper()
    if cycle == 'P':
        iloop = 7
        subcode = 46
if (semc == '3' or '4') and dip == 'Y':
    iloop = 9
    subcode = 58

# Opens file for storing data
with open('test2.txt', 'w+') as f:
    # print("   USN\t\t15MAT21\t\t15CHE22\t\t15PCD23\t\t15CED24\t\t15ELN25\t\t15CPL26\t\t15CHEL27\t15CIV28\n")

    c = 0
    pf = ''
    # For Loop to loop through all USNs
    for x in college:
        for u in range(low, high):

            # IF condition to concatenate USN
            if u < 10:
                usn = x + year + branch + '00' + str(u)
            elif u < 100:
                usn = x + year + branch + '0' + str(u)
            else:
                usn = x + year + branch + str(u)

            # opens the vtu result login page, gets the usn and opens the result page
            url = "http://results.vtu.ac.in/vitaviresultcbcs/index.php"
            if semc == '7':
                url = "http://results.vtu.ac.in/vitaviresultnoncbcs/index.php"
            br = RoboBrowser()
            br.open(url)
            form = br.get_form()
            form['lns'].value = usn
            br.submit_form(form)
            soup = br.parsed

            # Finds all the table elements and stores in array tds
            tds = soup.findAll('td')
            ths = soup.findAll('th')
            divs = soup.findAll('div', attrs={'class': 'col-md-12'})
            divCell = soup.findAll('div', attrs={'class': 'divTableCell'})

            try:
                sem = divs[5].div.text
                sem = sem.strip('Semester : ')
            except AttributeError:
                print("INVALID USN/ INCOMPATIBLE DATA")

            # IF condition to check invalid page opener
            if tds[0].text != 'University Seat Number ' or sem != semc:  # To check for Diploma and Number of students
                print("INVALID USN/ INCOMPATIBLE DATA")
                continue
            record = ''

            # if c == 0:
            #     c += 1
            #     for i in range(6, subcode, 6):
            #         # sublist += divCell[i].text + " ,"
            #         record = record + divCell[i].text + ","
            #     record += "\n"

            # tds[1] holds USN number
            record += re.sub('[!@#$:]', '', tds[1].text)
            record += ','
            # tds[3] holds the name
            record += re.sub('[!@#$:]', '', tds[3].text)
            record += ','

            sortList1 = []
            for i in range(6, subcode, 6):
                if (divCell[i].text[-3:]).isdigit():
                    sortList1.append(divCell[i].text[-3:])
                else:
                    sortList1.append(divCell[i].text[-2:])
            sortList1.sort()

            # for i in range(0,8):
            #     for j in range(6, subcode, 6):
            #

            ilist = []
            for i in range(0, iloop):
                for j in range(6, subcode, 6):
                    if (divCell[j].text[-3:]).isdigit():
                        if divCell[j].text[-3:] == sortList1[i] and j not in ilist:
                            ilist.append(j)
                    else:
                        if divCell[j].text[-2:] == sortList1[i] and j not in ilist:
                            ilist.append(j)

            # Strips extra garbage from the retrieved USN text
            print(record, end='\t')
            # Loop that goes from 8 to 51 in steps of 6 because starting from 8, in steps of 6
            try:
                for l in ilist:
                    # Checks if string has number
                    for j in range(l, l + 6):
                        if j == l + 1:
                            continue
                        else:
                            char = divCell[j].text
                            if char.isdigit():
                                record = record + str(int(char)) + ','
                            else:
                                record = record + char + ','
                            print(divCell[j].text, end='\t\t')
                            if j == l + 5:
                                pf = pf + divCell[j].text + ','
                # Writes the record into the file
                if semc == '7':
                    record += re.sub('[!@#$:]', '', tds[5].text)
                    print(re.sub('[!@#$:]', '', tds[5].text), end='\t\t')
                    record += ','
                    record += re.sub('[!@#$:]', '', tds[7].text)
                    print(re.sub('[!@#$:]', '', tds[7].text), end='\t\t')
                    record += ','
                f.write(record + '\n')
                print('\n')
            except IndexError:
                print("INVALID USN/ INCOMPATIBLE DATA")

if semc != '7' and dip != 'Y':
    if semc != '1':
        cycle = 'N'
    from sgparank import gpa2

    gpa2(year, branch, low, high, sem, cycle)
