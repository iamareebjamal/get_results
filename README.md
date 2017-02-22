# AMU B.Tech Results Downloader Python Script

[![Codacy grade](https://img.shields.io/codacy/grade/5a696561ef25409fa0531459e169eb8d.svg)]()
[![Website](https://img.shields.io/website-up-down-green-red/http/ctengg.amu.ac.in.svg)]()
[![License: GPL v3](https://img.shields.io/badge/License-GPL%20v3-blue.svg)](http://www.gnu.org/licenses/gpl-3.0)  

> Python Script to download results of whole class/branch by providing attendance Excel file. :wink:

##Description 
This script created by iamareebjamal batch downloads the result of B.Tech students hosted on http://ctengg.amu.ac.in by reading student information from Attendance Excel files. 

This script is provided on as-is basis and there is no responsibility of the developer to ensure proper working and genuinety of results downloaded. *Use on your risk.* :innocent:

Code of the program is open *(obviously, Python)* and it is encouraged to use it for educational purposes, tweaking, merging with other programs **strictly** till it remains personal. If intended otherwise, it should follow the GNU LICENSE provided. 

There's no guarantee that code will be easy to understand. Program is modular and does similar works in different modules so redundancy is natural. Assume it to be dirty coded. Be my guest to optimise it. Don't expect helpful comments. :stuck_out_tongue:

######Requirements
+ Requires Python 3 Environment set up, (duh). :unamused:
+ Needs these Python Modules to be installed - `BeautifulSoup`, `requests`, `xlsxwriter`, `xlrd`. Install libraries using `pip install -module_name-`. 
+ ~~Currently, script is made to run on Android, but will only require change of directories in `enlist.py` and `utils.py` for working in Windows, Linux, and OS X.~~
+ Works on Windows, Linux and Mac. Although script is tested briefly on Windows only. 
+ This branch hosts Windows, Linux and Mac script while Android one will be found in `android` branch. 
+ Requires internet connection to download results. Option 2 and 3 *(Read [Instructions](#instructions) in `main.py`)* do not require internet if result has been downloaded previously. 
+ Requires an Attendance Excel file to access student name, faculty number and enrolment number. 

######To-Do
- [x] Create android branch.
- [x] Fix master branch for Windows, Linux, Mac.
- [x] Fix and generalize directories and test.

######Screenshots
![downloading results](https://raw.githubusercontent.com/iamareebjamal/get_results/master/pics/pic1.png)
![downloaded results](https://raw.githubusercontent.com/iamareebjamal/get_results/master/pics/pic2.png)
![created spreadsheet](https://raw.githubusercontent.com/iamareebjamal/get_results/master/pics/pic3.png)

######Functions 
The program does 3 tasks :

1. Reads information from Attendance Excel file and downloads the result from site and saves in `/iaj/Store/` as HTML file.
2. Iterates through downloaded HTML and loads CPI and SPI of students.
3. Creates an Excel file with Student information and their CPI and SPI.

####Instructions
1. Place all 3 `.py` files in any folder in which you have writing permissions. Like `D:\Results\` or `~/home/user/results/`. *Normally, `C:\` drive on Windows doesn't have writing permissions.
2. Run the main script, `main.py` so it creates all necessary folders and read the instructions it prints. 
3. Now, script will demand a file name. You need to put the Attendance Excel file in `/Input/` folder in order to do so. 
4. Now give the exact name of Excel file to the script. *Eg:* `AM111.xlsx`
5. Now, the script will present 3 options if file is loaded successfully:
  1. Download Results. 
  2. Load Marks. 
  3. Create Worksheet. 
Choose relevant option based on [functions](#functions) listed above or the instructions printed by the script.
6. Downloaded results are stored in `/iaj/Store`. You can check individual results in any HTML Viewer or Browser. 
7. The worksheet is created in `/Output/` folder and contains CPI and SPI along with student information. 

**Note** : The script saves some configuration and database files in `/iaj/` folder for future faster loading and performance purposes. Please don't delete them :sweat:. You can delete the HTML files in Store folder to download result again. :+1:

Happy 'Result'ing and best of luck! :dancers:

All Hail Python :raised_hands:
