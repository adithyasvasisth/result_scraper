import sys
import re
import warnings
from robobrowser import RoboBrowser
from bs4 import BeautifulSoup as soup

if not sys.warnoptions:
    warnings.simplefilter("ignore")


def num_there(s):
    return any(i.isdigit() for i in s)


# List of Subjects
subj = ['1', '2', '3', '4', '5', '6', '7', '8']
ch = 0

# Input for Branch and USNs

college = input("Enter the college code\n").upper()
year = input('Enter the year\n')
branch = input('Please enter the branch\n').upper()
low = int(input('Enter starting USN\n'))
# increment last USN to aid looping
high = int(input('Enter last USN\n')) + 1
semc = input('Enter the Semester\n')
cycle = input('Enter the Cycle\n').upper()
if semc == '1' and cycle == 'P':
    subcode = 46
    markscode = 50
else:
    subcode = 52
    markscode = 56

# Opens file for storing data
with open('test.txt', 'w+') as f:
    # print("   USN\t\t15MAT21\t\t15CHE22\t\t15PCD23\t\t15CED24\t\t15ELN25\t\t15CPL26\t\t15CHEL27\t15CIV28\n")

    c = 0
    pf = ''
    # For Loop to loop through all USNs
    for i in range(low, high):

        # IF condition to concatenate USN
        if i < 10:
            usn = '1' + college + year + branch + '00' + str(i)
        elif i < 100:
            usn = '1' + college + year + branch + '0' + str(i)
        else:
            usn = '1' + college + year + branch + str(i)

        # opens the vtu result login page, gets the usn and opens the result page
        br = RoboBrowser()
        br.open("http://results.vtu.ac.in/vitaviresultcbcs/index.php")
        form = br.get_form()
        form['usn'].value = usn
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

        if c == 0:
            c += 1
            for i in range(6, subcode, 6):
                record = record + divCell[i].text + ","
            record += "\n"

        # print (ths[0].text)
        # tds[1] holds USN number
        record += re.sub('[!@#$:]', '', tds[1].text)
        record += ','
        # tds[3] holds the name
        record += re.sub('[!@#$:]', '', tds[3].text)
        record += ','

        # Strips extra garbage from the retrieved USN text

        print(record, end='\t')

        # Loop that goes from 8 to 51 in steps of 6 because starting from 8, in steps of 6
        for l in range(8, markscode, 6):
            # Checks if string has number
            for j in range(l, l + 4):
                record = record + divCell[j].text + ','
                print(divCell[j].text, end='\t\t')
                if j == l + 3:
                    pf = pf + divCell[j].text + ','
        # Writes the record into the file
        f.write(record + '\n')
        print('\n')

        # Loop to read data from file and converting marks to int and calculating highest in each subject

    # f.seek(0)
    # spn = [0, 0, 0, 0, 0, 0, 0, 0]
    # sn = [0, 0, 0, 0, 0, 0, 0, 0]
    # maxmks = [0, 0, 0, 0, 0, 0, 0, 0]
    # maxmksusn = ['', '', '', '', '', '', '', '']
    # k = 0
    # for i in range(low, high):
    #     record = f.readline()
    #     usn = record[0:10]
    #     record = record[11:]
    #     for j in range(0, 8):
    #         mark = ''
    #         k = 0
    #         if record == '':
    #             break
    #         while k < len(record):
    #             mark = mark + record[k]
    #             k = k + 1
    #             if (record[k] == '\n' or record[k] == '\t'):
    #                 break
    #         record = record[(k + 1):]
    #         if (maxmks[j] < int(mark)):
    #             # print(mark)
    #             maxmks[j] = int(mark)
    #             maxmksusn[j] = usn
    #         if (int(mark) >= 90):
    #             spn[j] += 1
    #         elif (int(mark) >= 80):
    #             sn[j] += 1
    #
    # for i in range(0, 8):
    #     print('The student with max marks in ' + subj[i] + ' is ' + maxmksusn[i] + ' with marks ' + str(maxmks[i]))
    #     print('')
    #     if i > 6:
    #         continue
    #     print('Number of S+ students: ' + str(spn[i]))
    #     print('Number of S students: ' + str(sn[i]))
    #     print('\n')

import xlwt

book = xlwt.Workbook()
ws = book.add_sheet('Sheet1')  # Add a sheet

f = open('test.txt', 'r+')

alignment = xlwt.Alignment()  # Create Alignment
font = xlwt.Font()  # Create the Font
alignment.horz = xlwt.Alignment.HORZ_CENTER
font.bold = True
style = xlwt.XFStyle()  # Create Style
# style.alignment = alignment  # Add Alignment to Style

data = f.readlines()  # read all lines at once
for i in range(len(data)):
    row = data[i].split(',')
    # This will return a line of string data, you may need to convert to other formats depending on your use case
    if i == 0:
        k = 1
        ws.write(i, 0, "USN")
        ws.write(i, 1, "NAME")
        for j in range(len(row)):
            ws.write_merge(i, i, k + 1, k + 4, row[j], style)
            k += 4
    else:
        for j in range(len(row)):
            ws.write(i, j, row[j], style)  # Write to cell i, j

pth = 'ExcelFiles/'
book.save(pth + '1' + college + year + branch + '.xlsx')

f.close()

from sgpa import gpa

gpa(college, year, branch, sem, cycle)

# from subj import subres

# subres(college, year, branch, sem, pf)
