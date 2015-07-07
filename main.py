import requests 
import os
import enlist
import time 
import utils
from bs4 import BeautifulSoup 

os.chdir('/sdcard/Project/')

url = 'http://ctengg.amu.ac.in/result_btech.php'
    

def for_student(fac_no, en_no, name):
  r_no = fac_no[5:8]
  path = '/sdcard/Project/Store/'
  file_name = r_no + ' - ' + name + ' (' + fac_no + ')' +'.html'
  os.chdir(path)
  if os.path.isfile(file_name):
    print('File Exists... Skipping\n\t', file_name) 
  else :
    try:
      form_data = {'FN':fac_no, 'EN':en_no, 'submit':'submit'}
      response = requests.post(url, data=form_data)
      if 'CPI' in response.text:
        soup = BeautifulSoup(response.text)
        with open(file_name , 'w+') as ou:
          print(soup.prettify(), file=ou)
        print('Saved result of', name)
      else :
        print('Wrong input data or no result...') 
        #print('Doesnt exist', file_name)
    except requests.exceptions.ConnectionError as err:
      print('No Connection')
  
def get_result(students):
  for key in students.keys():
    name = students[key]['name'] 
    en_no = students[key]['en_no']
    fac_no = students[key]['fac_no']
    print('\nGetting result of ', fac_no, name)
    for_student(fac_no, en_no, name)
    #time.sleep(.5)
  print('\n\nAll results saved')
    


def main():
  name = 'store.xlsx'
  
  students = enlist.populate(name)

  if not students == None :
    get_result(students)
    utils.set_marks(students)
    print(utils.create_worksheet(name, students))
  else:
    print('Error reading student database worksheet')


# Let's Go! 
main() 