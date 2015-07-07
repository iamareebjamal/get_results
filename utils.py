from bs4 import BeautifulSoup 
import os
import ast
import xlsxwriter

os.chdir('/sdcard/Project/') 

def save(students):
  os.chdir('/sdcard/Project/') 
  with open('results.db', 'w+') as fi:
      print(students, file=fi)

def load_all(students):
  for key in students.keys():
    name = students[key]['name']
    fac_no = students[key]['fac_no']
    r_no = fac_no[5:8]
    path = '/sdcard/Project/Store/'
    file_name = r_no + ' - ' + name + ' (' + fac_no + ')' +'.html'
    os.chdir(path)
    with open(file_name, 'r+') as fi:
      page = fi.read()
      soup = BeautifulSoup(page)
    try:
      spi = soup('table')[2].findAll('tr')[1].findAll('th')[5].string.strip()
      cpi = soup('table')[2].findAll('tr')[1].findAll('th')[4].string.strip()
    except IndexError as err:
      spi = 0
      cpi = 0
    students[key]['spi'] = spi
    students[key]['cpi'] = cpi
    print(students[key]['name'], 'done') 
  save(students)
  return students

def load(students, key):
  name = students[key]['name']
  fac_no = students[key]['fac_no']
  r_no = fac_no[5:8]
  path = '/sdcard/Project/Store/'
  file_name = r_no + ' - ' + name + ' (' + fac_no + ')' +'.html'
  os.chdir(path)
  with open(file_name, 'r+') as fi:
    page = fi.read()
    soup = BeautifulSoup(page)
  try:
    spi = soup('table')[2].findAll('tr')[1].findAll('th')[5].string.strip()
    cpi = soup('table')[2].findAll('tr')[1].findAll('th')[4].string.strip()
  except IndexError as err:
    spi = 0
    cpi = 0
  return {'cpi':cpi, 'spi' : spi}
    
def set_marks(students):
  os.chdir('/sdcard/Project/') 
  if os.path.isfile('results.db'):
    with open('results.db', 'r+') as fi:
      data = fi.read() 
      res = ast.literal_eval(data)
    for key in students.keys():
      try:
        students[key]['cpi'] = res[key]['cpi']
        students[key]['spi'] = res[key]['spi']
      except KeyError as err :
        students[key]['cpi'] = load(students, key)['cpi']
        students[key]['spi'] = load(students, key)['spi']
    save(students)
    return students
  else:
    load_all(students)
