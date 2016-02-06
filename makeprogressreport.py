import datetime
import os.path
import os
import sys
from shutil import copyfile
from openpyxl import load_workbook
import time

print('Welcome to makeProgressReport!')
time.sleep(1)
print('------------------------------')

#get user name if this is the first time the project has been run
#if the userData.txt file doesn't exist, ask for the users name and create
#the userData.txt file
if (os.path.isfile('userData.txt')!= True):
   dataFile = open('userData.txt','a')
   print('It seems like this is the first time \nyou have used makeProgressReport!')
   time.sleep(1)
   print('...')
   time.sleep(1)
   print('Please enter your first name.')
   firstName = input()
   firstName = firstName.strip()
   firstName = firstName.title()
   time.sleep(1)
   print('Please enter your last name.')
   lastName = input()
   lastName = lastName.strip()
   lastName = lastName.title()
   time.sleep(1)
   #make the user verify her name
   print('is this your name: '+firstName+' '+lastName+'? (y/n)')
   answer = str(input())
   answer = answer.lower()
   answer = answer.strip()
   if(answer == 'y'):
      #if they verify, continue the rest of the program
      print('Hello, '+firstName+'!')
      time.sleep(1)
      print('saving your data to userData.txt')
      time.sleep(1)
      dataFile.write(firstName+'\n'+lastName)
      dataFile.close()
   else:
      #if they mess up their name, have them restart from the beginning
      print('come on man...close the terminal and try again')
      dataFile.close()
      os.remove('userData.txt')
      input()
      sys.exit()
else:
   #if they've created their user profile already
   #get the user information from the file
   dataFile = open('userData.txt')
   firstName = dataFile.readline().rstrip()
   lastName = dataFile.readline().rstrip()
   print('Hello, '+firstName+'!')
   dataFile.close()

   
#template information
dateCell = 'B7'
nameCell = 'A6'
templateName = 'Progress_Template'
extension = '.xlsx'

#figure out what last saturday's date was
today = datetime.date.today()
#if today is saturday, it is the start of the week
if (today.weekday() == 5):
   sat = today
else:
   #if today is not satufday, find out what the last saturday that happened was
   idx = (today.weekday()+1)%7
   sat = today - datetime.timedelta(7+idx-6)

#create file name from previously collected saturday information
fileName = str(sat.year)+str(sat.month).zfill(2)+str(sat.day).zfill(2)+firstName+lastName+'Progress'
fileName = fileName[2:]

#check if the file already exists (ie we created it already)
if os.path.isfile(fileName+extension):
   print('I, have already created a progress \nreport for this week in this directory.')
   input()
   sys.exit()

else:
    #if it doesn't exist make it
    #make a copy of the template project in the directory
    copyfile(templateName+extension,fileName+extension)
    #open the workbook
    progressWorkbook = load_workbook(fileName+extension)
    #get the first sheet
    firstWorksheet = progressWorkbook.active
    #set the date cell correctly
    firstWorksheet[dateCell] = sat.strftime("%B %d, %Y")
    #edit the name cell
    firstWorksheet[nameCell] = firstName+' '+lastName
    #save the file
    progressWorkbook.save(fileName+extension)
    print('All Finished, See you next week!')
    input()
    sys.exit()



    
