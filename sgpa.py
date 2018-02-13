# coding=utf-8

def grade(m):
    grade_point = 0
    grade_letter = ''
    if m >= 90:
        grade_letter += 'S+,'
        grade_point = 10
    elif m >= 80:
        grade_letter += 'S,'
        grade_point = 9
    elif m >= 70:
        grade_letter += 'A,'
        grade_point = 8
    elif m >= 60:
        grade_letter += 'B,'
        grade_point = 7
    elif m >= 50:
        grade_letter += 'C,'
        grade_point = 6
    elif m >= 45:
        grade_letter += 'D,'
        grade_point = 5
    elif m >= 40:
        grade_letter += 'E,'
        grade_point = 4
    elif m <= 39:
        grade_letter += 'F,'
        grade_point = 0
    return grade_letter, grade_point


def calc(sub, count1, count2, record):
    c = count1 * 4 + count2 * 2
    cp = 0
    for i in range(0, count1):
        st, g = grade(sub.pop(0))
        record += st
        cp += g * 4
    for i in range(0, count2):
        st, g = grade(sub.pop(0))
        record += st
        cp += g * 2
    gp = str(round((cp / c), 2))
    return record, gp


def gpa(college, year, branch, sem, cycle):
    if int(sem) == 1:
        count1 = 5
        count2 = 2
    else:
        count1 = 6
        count2 = 2
    if cycle == 'P':
        marks_code = 29
        gpa_col = 9
    else:
        marks_code = 33
        gpa_col = 10
    pth = 'ExcelFiles/'
    import xlrd
    wb = xlrd.open_workbook(pth + '1' + college + year + branch + '.xlsx')
    sheet = wb.sheet_by_name('Sheet1')
    sub = []
    with open('gpa.txt', 'w+') as f:
        record = ''
        record += sheet.cell_value(0, 0) + ',' + sheet.cell_value(0, 1) + ','
        print(record, end='\t')
        f.write(record)
        for i in range(2, sheet.ncols - 1, 4):
            record = ''
            record += sheet.cell_value(0, i) + ','
            print(record, end='\t')
            f.write(record)
        print('\n')
        f.write('\n')
        for i in range(1, sheet.nrows):
            record = ''
            record += sheet.cell_value(i, 0) + ',' + sheet.cell_value(i, 1) + ','
            for j in range(4, marks_code, 4):
                if sheet.cell_value(i, j + 1) == 'P':
                    sub.append(int(sheet.cell_value(i, j)))
                else:
                    sub.append(0)
            record, sgpa = calc(sub, count1, count2, record)
            percent = str((float(sgpa) - 0.750) * 10)
            record += sgpa + ',' + percent + ','
            print(record, end='\t')
            print('\n')
            f.write(record + '\n')
    f.close()

    import xlwt

    book = xlwt.Workbook()
    ws = book.add_sheet('Sheet1')
    f = open('gpa.txt', 'r+')
    data = f.readlines()  # read all lines at once
    for i in range(len(data)):
        row = data[i].split(',')
        for j in range(len(row)):
            if i == 0 and j == gpa_col:
                pass
            else:
                ws.write(i, j, row[j])  # Write to cell i, j
    ws.write(0, gpa_col, 'SGPA')
    ws.write(0, gpa_col + 1, 'PERCENTAGE')

    book.save(pth + '1' + college + year + branch + 'GPA.xlsx')

    f.close()
