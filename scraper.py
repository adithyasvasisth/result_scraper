import sys
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

college = input("Enter the college code\n")
year = input('Enter the year\n')
branch = input('Please enter the branch\n')
low = int(input('Enter starting USN\n'))
# increment last USN to aid looping
high = int(input('Enter last USN\n')) + 1
semc = input('Enter the Semester\n')
# print('Diploma student?(y/n)')
# dpl = input()
# if (dpl == 'y'):
#     ch = 57
# else:
#     ch = 51

# Opens file for storing data
with open('test.txt', 'w+') as f:
    # print("   USN\t\t15MAT21\t\t15CHE22\t\t15PCD23\t\t15CED24\t\t15ELN25\t\t15CPL26\t\t15CHEL27\t15CIV28\n")

    c = 0

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
        if tds[0].text != 'University Seat Number ' or sem != semc:  # To check for Diploma and Number of students : (len(tds) < ch or len(ths) < 6)
            print("INVALID USN/ INCOMPATIBLE DATA")
            continue
        record = ''

        if c == 0:
            c += 1
            record = ""
            record += "\t\t"
            for i in range(6, 46, 6):
                record = record + divCell[i].text + "\t\t"
            record += "\n"

        # print (ths[0].text)
        # tds[1] holds USN number
        record += tds[1].text

        # Strips extra garbage from the retrieved USN text
        record = record.strip(' : ')

        print(record, end='\t')

        # Loop that goes from 8 to 51 in steps of 6 because starting from 8, in steps of 6, you get final marks of each subject
        for j in range(10, 52, 6):
            # Checks if string has number
            # for j in range(l, l+3, 1):
            if num_there(divCell[j]):
                record = record + '\t' + divCell[j].text
                print(divCell[j].text, end='\t\t')
            else:
                break
            print('')

        # Writes the record into the file
        f.write(record + '\n')

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
import xlrd

book = xlwt.Workbook()
ws = book.add_sheet('First Sheet')  # Add a sheet

f = open('test.txt', 'r+')

data = f.readlines()  # read all lines at once
for i in range(len(data)):
    row = data[
        i].split()  # This will return a line of string data, you may need to convert to other formats depending on your use case

    for j in range(len(row)):
        ws.write(i, j, row[j])  # Write to cell i, j

book.save('Excelfile' + '.xls')
f.close()

input()
