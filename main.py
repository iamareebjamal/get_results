import requests 
import os
import enlist
import time 
import utils
from bs4 import BeautifulSoup 

welcome = \
"""

AMU B.Tech Results Downloader
This Python script downloads B.Tech Results of whole class based on information in attendance Excel file.

First you need to put Excel file in Input/ folder
Then type the name of file when asked (Default : store.xlsx) 
This will load the information about students from the Excel file and stores it in students.db for future faster access.
Then you will prompted 3 options:

    First option downloads the result of whole class and 
    stores them as html pages in Store/ folder. 
    Note : This option should be run at least once to 
    download  all necessary result files for further options
    If there are no result files in Store/ folder, then 
    script will not run properly. 
   
    Second option loads CPI and SPI from downloaded 
    html to script database as results.db for future faster
    access. 
    Note : This option is necessary to be run in 
    order to run 3rd option. If this option is not run, 
    then no data can be written in Excel file. 
    
    Third option reads your CPI and SPI from Updated 
    database and saves the information as Excel file in 
    Output/ folder.
    
    
Note: App creates required files and databases in iaj/ folder for proper functioning. Please don't delete those files. 
    
    
Let's start:"""

os.chdir('/sdcard/Project/')

url = 'http://ctengg.amu.ac.in/result_btech.php'
    
# Let's Go!


print(welcome) 

def for_student(fac_no, en_no, name):
  r_no = fac_no[5:8]
  path = '/sdcard/Project/iaj/Store/'
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
    


def main(students):
  menu = \
'''


1. Download Results. 
2. Load Marks.
3. Create result worksheet.
4. Exit

'''
  
  print(menu)
  x = input()
  
  if x == '1' :
    get_result(students)
    input('Press any key to Continue...')
  elif x == '2' :
    utils.set_marks(students)
    input('Press any key to Continue...')
  elif x == '3' : 
    print(utils.create_worksheet(name, students))
    input('Press any key to Continue...')
  elif x == '4' :
    return
  else:
    print('Please choose valid option')
    input('\n\n\nPress any key to Continue...')
  main(students)

wrong=1
while(wrong):
  utils.mkdirs()
  name = input('Enter the Excel file name: ').rstrip()
  students = None
  students = enlist.populate(name)
  if students == None:
    print('Error reading student database...\nRetry running script')
  else:
    wrong = 0
    main(students)