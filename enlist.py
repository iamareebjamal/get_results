import xlrd
import os
import ast
import utils


def cd(a=0):
    if a == 0:
        path = os.path.join(utils.path(), 'Input')
    elif a == 1:
        path = os.path.join(utils.path(), 'iaj')
    os.chdir(path)

file_names = dict()
def list_input():
    cd()
    file_names.clear()
    counter = 1
    for li in os.listdir():
        file_names[counter] = li 
        print(str(counter) + ". " + li)
        counter += 1
         

def is_fac_no(fac):
    try:
        (start, finish) = fac[0:2], fac[5:8]
        return start.isdigit() and finish.isdigit() and len(fac) == 8
    except TypeError:
        return False


def find_index(sheet):
    for ro in range(0, sheet.nrows):
        for co in range(0, sheet.ncols):
            if is_fac_no(sheet.cell_value(rowx=ro, colx=co)):
                return [ro, co]


def file_changed(file_name):
    cd(1)
    if not os.path.isfile('conf.ini'):
        with open('conf.ini', 'w') as fi:
            print(file_name, file=fi)
        with open('students.db', 'w+') as fi:
            print('', file=fi)
        return True
    else:
        with open('conf.ini', 'r') as fi:
            old = fi.read().rstrip()
        if old == file_name:
            return False
        else:
            with open('conf.ini', 'w+') as fi:
                print(file_name, file=fi)
            return True


def is_worksheet(name):
    cd()
    if os.path.isfile(name):
        try:
            xlrd.open_workbook(name)
        except xlrd.biffh.XLRDError:
            print('Unsupported Format or corrupt file')
            return False
        return True
    else:
        return False


def get_students(file_name):
    cd()
    if is_worksheet(file_name):
        store = xlrd.open_workbook(file_name)
        sheet = store.sheet_by_index(0)
        i = 1
        j, k = find_index(sheet)[0], find_index(sheet)[1]
        index = range(j, sheet.nrows)
        students = dict()

        for ro in index:
            fac_no = sheet.cell_value(rowx=ro, colx=k)
            en_no = sheet.cell_value(rowx=ro, colx=k + 1)
            name = sheet.cell_value(rowx=ro, colx=k + 2).replace('_', ' ')
            students[i] = dict()
            students[i]['fac_no'] = fac_no
            students[i]['en_no'] = en_no
            students[i]['name'] = name
            i = i + 1
        cd(1)
        with open('students.db', 'w+') as fi:
            print(students, file=fi)
        return students
    else:
        return None


def populate(counter, reset=0):
    path = os.path.join(utils.path(), 'iaj', 'Store')
    if not os.path.exists(path):
        os.makedirs(path)
    cd(1)
    file_name = file_names[counter]
    if os.path.isfile('students.db') and not file_changed(file_name) and not reset:
        with open('students.db', 'r+') as fi:
            s = fi.read()
            students = ast.literal_eval(s)
        try:
            if students[1]['name'] in s:
                return students
            else:
                print('Bad dictionary file...\nRetrying reading from Excel file')
                return get_students(file_name)
        except KeyError:
            print('Bad dictionary file... \nRetrying reading Excel file')
            return get_students(file_name)
    else:
        return get_students(file_name)
