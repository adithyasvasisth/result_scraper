# coding=utf-8

def grade(m):
    if m >= 91:
        return 10
    elif m >= 81:
        return 9
    elif m >= 71:
        return 8
    elif m >= 61:
        return 7
    elif m >= 51:
        return 6
    elif m >= 45:
        return 5
    elif m >= 40:
        return 4
    elif m <= 39:
        return 0


def calc(sub, count1, count2):
    c = count1*4 + count2*2
    cp = 0
    for i in range(0, count1):
        cp += grade(sub.pop(0)) * 4
    for i in range(0, count2):
        cp += grade(sub.pop(0)) * 2
    return cp/c


def gpa():
    sub = []
    sem = int(input("Which semester ? "))
    if sem == 1:
        count1 = 5
        count2 = 2
    else:
        count1 = 6
        count2 = 2
    for i in range(1, count1 + 1):
        sub.append(int(input("Subject %d marks ? " % i)))
    for i in range(count1 + 1, count2 + count1 + 1):
        sub.append(int(input("Subject %d marks ? " % i)))
    sgpa = str(round(calc(sub, count1, count2), 2))
    print("\n--------------------\n The SGPA is : " + sgpa + "\n--------------------\n")

gpa()