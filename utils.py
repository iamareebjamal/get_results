from bs4 import BeautifulSoup
import os
import ast
import xlsxwriter

def path():
  return os.path.dirname(os.path.realpath(__file__))

def mkdirs():
  dirs = [os.path.join(path(), 'iaj', 'Store'), os.path.join(path(), 'Input'), os.path.join(path(), 'Output')]
  for dire in dirs :
    if not os.path.exists(dire):
      os.makedirs(dire)

def cd(a=0):
  if a == 0:
    pathe = os.path.join(path(), 'iaj')
  elif a == 1:
    pathe = os.path.join(path(), 'Output')
  elif a == 2:
    pathe = os.path.join(path(),'iaj', 'Store')
  os.chdir(pathe)
  


def save(students):
  cd()
  with open('results.db', 'w+') as fi:
      print(students, file=fi)

def load_all(students):
  for key in students.keys():
    name = students[key]['name']
    fac_no = students[key]['fac_no']
    r_no = fac_no[5:8]
    file_name = r_no + ' - ' + name + ' (' + fac_no + ')' +'.html'
    cd(2)
    try:
      with open(file_name, 'r+') as fi:
        page = fi.read()
        soup = BeautifulSoup(page)
    except IOError:
      print('HTML file not found. Download files by option 1 first...') 
      return
    try:
      spi = soup('table')[1].findAll('td')[4].string.strip()
      cpi = soup('table')[1].findAll('td')[5].string.strip()
    except IndexError:
      spi = 0
      cpi = 0
    students[key]['spi'] = spi
    students[key]['cpi'] = cpi
    print('Loaded:', students[key]['name'])
  save(students)
  return students

def load(students, key):
  name = students[key]['name']
  fac_no = students[key]['fac_no']
  r_no = fac_no[5:8]
  file_name = r_no + ' - ' + name + ' (' + fac_no + ')' +'.html'
  cd(2)
  with open(file_name, 'r+') as fi:
    page = fi.read()
    soup = BeautifulSoup(page)
  try:
    spi = soup('table')[2].findAll('tr')[1].findAll('th')[4].string.strip()
    cpi = soup('table')[2].findAll('tr')[1].findAll('th')[5].string.strip()
  except IndexError:
    spi = 0
    cpi = 0
  return {'cpi':cpi, 'spi' : spi}
    
def set_marks(students):
  cd()
  if os.path.isfile('results.db'):
    with open('results.db', 'r+') as fi:
      data = fi.read()
      res = ast.literal_eval(data)
    if students[1]['name']==res[1]['name']:
      pass
    else :
      print('Database mismatch. Did you download result first?')
      return
    for key in students.keys():
      try:
        students[key]['cpi'] = res[key]['cpi']
        students[key]['spi'] = res[key]['spi']
      except KeyError :
        students[key]['cpi'] = load(students, key)['cpi']
        students[key]['spi'] = load(students, key)['spi']
    save(students)
    print('Marks loaded')
    return students
  else:
    load_all(students)

def create_worksheet(name, students):
  pathe = os.path.join(path(), 'Output')
  if not os.path.exists(pathe):
    os.makedirs(pathe)
  cd(1)
  if os.path.isfile(name):
    over = input('Worksheet already exists. Do you want to over write?\nY for yes, any key for no... \n')
    if over == 'Y' or over == 'y':
      print('Over writing file') 
    else:
      return 'Skipping writing file...\n Check in Output/ directory'
  try :
    students[1]['cpi']
  except KeyError as err :
    return 'Wrong input file. Run 2. Load marks first...'
  data = xlsxwriter.Workbook(name)
  data_sheet = data.add_worksheet() 
  
  head = ['S. No.', 'Faculty No.', 'Enrolment No.', 'Name', 'CPI', 'SPI']
  col=0
  for cell in head:
    data_sheet.write(0,col, cell)
    col+=1

  for key in students.keys():
    col=0
    student = [key, students[key]['fac_no'], students[key]['en_no'], students[key]['name'], students[key]['cpi'], students[key]['spi']]
    for each in student:
      data_sheet.write(key, col, each)
      col+=1
      	
  data.close()
  success = 'Successfully created '+name+' in Output/ directory'
  return success
      
