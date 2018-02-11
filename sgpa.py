# coding=utf-8

def grade(m):
    if m >= 90:
        return 10
    elif m >= 80:
        return 9
    elif m >= 70:
        return 8
    elif m >= 60:
        return 7
    elif m >= 50:
        return 6
    elif m >= 45:
        return 5
    elif m >= 40:
        return 4
    elif m <= 39:
        return 0


def calc(sub, count1, count2):
    c = count1 * 4 + count2 * 2
    cp = 0
    for i in range(0, count1):
        cp += grade(sub.pop(0)) * 4
    for i in range(0, count2):
        cp += grade(sub.pop(0)) * 2
    return cp / c


def gpa(college, year, branch, sem):
    if int(sem) == 1:
        count1 = 5
        count2 = 2
    else:
        count1 = 6
        count2 = 2

    pth = 'ExcelFiles/'
    import xlrd
    wb = xlrd.open_workbook(pth + '1' + college + year + branch + '.xlsx')
    sheet = wb.sheet_by_name('Sheet1')
    sub = []
    with open('gpa.txt', 'w+') as f:
        for i in range(1, sheet.nrows):
            record = ''
            record += sheet.cell_value(i, 0) + ',' + sheet.cell_value(i, 1) + ','
            for j in range(4, 29, 4):
                sub.append(int(sheet.cell_value(i, j)))
            sgpa = str(round(calc(sub, count1, count2), 2))
            percent = str((float(sgpa) - 0.750) * 10)
            record += sgpa + ',' + percent + ','
            print(record, end='\t')
            print('\n')
            record.strip("'")
            f.write(record + '\n')

    f.close()

    import xlwt

    book = xlwt.Workbook()
    ws = book.add_sheet('Sheet1')
    f = open('gpa.txt', 'r+')
    data = f.readlines()  # read all lines at once
    ws.write(0, 0, 'USN')
    ws.write(0, 1, 'NAME')
    ws.write(0, 2, 'SGPA')
    ws.write(0, 3, 'PERCENTAGE')
    for i in range(len(data)):
        row = data[i].split(',')
        for j in range(len(row)):
            ws.write(i + 1, j, row[j])  # Write to cell i, j

    book.save(pth + '1' + college + year + branch + 'GPA.xlsx')

    f.close()
