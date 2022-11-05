#Import the necessary modules
import pandas, re, os

#This line of Pandas->inputoutput->format->file_type = Excel-> ExcelFormatter is used to format excel way of looking
#Header_Style determines the header row look and all the layout related to header
#Setting it none make header row look likes normal data row only
#It is done on specific requirement of appearance of Output, as specified by Sir
pandas.io.formats.excel.ExcelFormatter.header_style = None

#As pandas raise settingWithCopyWarning,
#As it occur due to chain assignment,
#so to avoid its display in terminal, i used options attribute which is used to configure and prevent it from displaying error and exception
pandas.options.mode.chained_assignment = None

#During execution, Pandas was raising a lot of FutureWarnings
#to Prevent them from being displayed
#importing module 'warnings' which provide a lot a functionalities to handle warnings
#simplefilter() is use to apply a certain action on specific type of warnings
#applied action of 'ignore' to 'FutureWanings', 'UserWarning' category
import warnings
warnings.simplefilter(action='ignore', category= (FutureWarning, UserWarning))

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

import csv
from random import randint
from time import sleep



def send_mail(fromaddr, frompasswd, toaddr, msg_subject, msg_body, file_path):
    try:
        msg = MIMEMultipart()
        print("[+] Message Object Created")
    except:
        print("[-] Error in Creating Message Object")
        return

    msg['From'] = fromaddr

    msg['To'] = toaddr

    msg['Subject'] = msg_subject

    body = msg_body

    msg.attach(MIMEText(body, 'plain'))

    filename = file_path
    attachment = open(filename, "rb")

    p = MIMEBase('application', 'octet-stream')

    p.set_payload((attachment).read())

    encoders.encode_base64(p)

    p.add_header('Content-Disposition', "attachment; filename= %s" % filename)

    try:
        msg.attach(p)
        print("[+] File Attached")
    except:
        print("[-] Error in Attaching file")
        return

    try:
        #s = smtplib.SMTP('smtp.gmail.com', 587)
        s = smtplib.SMTP('mail.iitp.ac.in', 587)
        print("[+] SMTP Session Created")
    except:
        print("[-] Error in creating SMTP session")
        return

    s.starttls()

    try:
        s.login(fromaddr, frompasswd)
        print("[+] Login Successful")
    except:
        print("[-] Login Failed")

    text = msg.as_string()

    try:
        s.sendmail(fromaddr, toaddr, text)
        print("[+] Mail Sent successfully")
    except:
        print('[-] Mail not sent')

    s.quit()


def isEmail(x):
    if ('@' in x) and ('.' in x):
        return True
    else:
        return False

FROM_ADDR = "changeme"
FROM_PASSWD = "changeme"
receiver = "mayank265@iitp.ac.in"
Subject = "Sending Attendance Report"
Body ='''
Respected Sir,

Please find attached consolidated attendance report of students of CS384 python course with this mail.   

Thanking You.

--
Yours Sincerely
2001CB26
'''

from datetime import datetime
start_time = datetime.now()

def attendance_report():
    #1 
    # Opening the input files
    f1 = pandas.read_csv('input_attendance.csv')
    f2 = pandas.read_csv('input_registered_students.csv')

    #Declaration of all the user-defined data structures, I will be using
    roll_to_name, mydict = dict(), dict()
    total_lectures,myset = set(),set()
    total_lectures_taken = dict()
    attendance_real, attendance_duplicate, attendance_invalid = dict(),dict(),dict()
    ''' 
    1. mydict is a simple dictionary which will be used in defining other structures, by making and suppling a copy of it
    2. myset is a set similar to mydict
    3. roll_to_name is a dictionary
        having key : roll number and value : name of this roll number

    4. total_lectures is a set containing valid dates on which lecture happened in form of string

    5. attendance_real is a dictionary of set
        key : Roll number
        value : a set which contains dates stored in form of string for each roll number when a person has actual attendance
        attendance_real = {'2001CB02':{'01-09-2022', '28-07-2022', .....}, '2001CB03':{'01-09-2022', '28-07-2022', .....}, .....}

    
    Now, attendance_duplicate, attendance_invalid, total_lectures_taken are dictionary of dictionary
    key : roll number
    value : another dictionary for which~
        key : date in form of string
        value : number of attendances
    example: 
    {'2001CB02': {'28-07-2022': 1, '01-08-2022': 1, .....},
     '2001CB03': {'28-07-2022': 1, '01-08-2022': 2, .....},
     .........}

    7. attendance_real has dates on which attendance is counted actual
    8. attendance_duplicate has dates on which attendance is counted as duplicate
    9. attendance_invalid has dates on which attendance is marked outside 2 to 3 PM
    '''

    #Created Regrex Expression for extracting Roll Number, Time, Date in between various values
    roll_number_pattern = re.compile(r'2001[A-Za-z]{2}[\d]{2}')
    time_pattern = re.compile(r'[\d]{2}:[\d]{2}')
    date_pattern = re.compile(r'[\d]{2}-[\d]{2}-[\d]{4}')


    #2 Initialisation of various data structures
    #iterrows is a function which iterrates all the rows
    #counter or variable at 1st place holds index (0 based indexing)
    #row holds entire row of csv file in form of string
    for counter, rows in f2.iterrows():
        #we can access any row of a particular column by
        #<file_pointer>['<column_label>'][counter] = <value>

        #inserting name in dictionary
        roll_to_name[f2['Roll No'][counter]]=f2['Name'][counter]        
        try:
            #initiallising with default dictionary so that python understand the type of data structure
            attendance_real[f2['Roll No'][counter]] = myset.copy()
            total_lectures_taken[f2['Roll No'][counter]] = {" ":0}
            attendance_duplicate[f2['Roll No'][counter]] = {" ":0}
            attendance_invalid[f2['Roll No'][counter]] = {" ": 0}
        except KeyError:
            print("KeyError in Part 2")
        except:
            print("Some other Error in Part 2")
    
    #Deleting the default empty key we inputted during initialisation
    for i in roll_to_name.keys():
        del attendance_duplicate[i][" "]
        del attendance_invalid[i][" "]
        del total_lectures_taken[i][" "]

    #3 Working
    for counter, rows in f1.iterrows():
        #This column contain null values, so to handle them
        if(f1['Attendance'][counter]!=""):
            #extracting roll number through slicing
            roll_number = str(f1['Attendance'][counter])[0:8]

            #Taking Timestamp in a variable
            Timestamp = f1['Timestamp'][counter]
            _time, _date = "",""

            try:
                #Using Regrex Expression, finding Time from the experssion using finditer()
                #it is use to find all patterns similar in the main string
                boom = re.finditer(time_pattern, Timestamp)
                #here boom will store beginning and ending index of all pattern occuring in the string
                for i in boom:
                    #This loop will run only ones as there is only one pattern which fits the Regrex
                    #obtaing time through slicing of indices
                    _time = (f'{Timestamp[i.start():i.end()]}')
            except IndexError:
                print("Index Error in Part 3")
            except:
                print("Some other error in Part 3")
            
            #Using Regrex Expression, finding Date from the experssion using finditer()
            #it is use to find all patterns similar in the main string
            boom = re.finditer(date_pattern, Timestamp)
            #here boom will store beginning and ending index of all pattern occuring in the string
            for i in boom:
                #This loop will run only ones as there is only one pattern which fits the Regrex
                #obtaing time through slicing of indices
                _date = (f'{Timestamp[i.start():i.end()]}')
            
            #Checking if the roll number is of a registered student
            if(roll_number in roll_to_name.keys()):
                
                #Pandas.Timestamp() is used pass string to pandas, so that it convert it into a Timestamp
                #Weekday() return day of the week on provided Timestamp
                #'0' for Monday and '3' for Thrusday                
                if(pandas.Timestamp(_date).weekday() in [0,3]):

                    #Adding this date to valid lecture occured
                    total_lectures.add(_date)

                    #Adding to  lecture count for this roll number
                    #Since KeyError was occuring, i adopted this method to add dictionary
                    #if the date is not present as key in dictionary, initiallise it with 1 value
                    if(_date not in attendance_duplicate[roll_number].keys()):
                        total_lectures_taken[roll_number].setdefault(_date, 1)
                    else :
                        #just increase the value
                        total_lectures_taken[roll_number][_date] += 1

                    #Extracting hour, minute through slicing
                    _Hour, _minute= int(_time[0:2]), int(_time[3:5])

                    #Checking if this attendance is actual or fake by comparing time
                    #Actual is in between 14:00:00 to 15:00:00
                    if( (_Hour==14 and (_minute>=0 and _minute<=59))  or (_Hour == 15 and _minute==0)):

                        #if attendance is real, just add it to the set of respective roll number
                        if(_date not in attendance_real[roll_number]):
                            attendance_real[roll_number].add(_date)
                        #if attendance is marked as real once, means remaining will go to duplicate
                        else:
                            #Since KeyError was occuring, i adopted this method to add dictionary
                            #if the date is not present as key in dictionary, initiallise it with 1 value
                            if(_date not in attendance_duplicate[roll_number].keys()):
                                attendance_duplicate[roll_number].setdefault(_date, 1)
                            else :
                                #just increase the count
                                attendance_duplicate[roll_number][_date] += 1

                    else :
                        #Attendance is invalid
                        #Since KeyError was occuring, i adopted this method to add dictionary
                        #if the date is not present as key in dictionary, initiallise it with 1 value
                        if(_date not in attendance_invalid[roll_number].keys()):
                            attendance_invalid[roll_number].setdefault(_date, 1)
                        else :
                            #just increase the value
                            attendance_invalid[roll_number][_date] += 1


    
    #4 Inserting data to individual reports for each roll number and consolidated report for all roll numbers in a single file
    #Creating data frame for consolidated report
    f3 = pandas.DataFrame()

    #Inserting columns in dataframe
    # <file Pointer>.insert(index, '<column_label>')
    f3.insert(len(f3.columns), 'Roll', '')
    f3.insert(len(f3.columns), 'Name', '')
    #Inserting dynamic date columns through for loop
    for i in total_lectures:
        f3.insert(len(f3.columns), i, '')
    f3.insert(len(f3.columns), 'Actual Lecture Taken', '')
    f3.insert(len(f3.columns), 'Total Real', '')
    f3.insert(len(f3.columns), '% Attendance', '')
    #variable to keep track of index of row
    index_f3 = 0

    for i in roll_to_name.keys():
        #new dataframe for individual reports
        f4 = pandas.DataFrame()

        #Inserting columns
        f4.insert(len(f4.columns), 'Date', '')
        f4.insert(len(f4.columns), 'Roll', '')
        f4.insert(len(f4.columns), 'Name', '')
        f4.insert(len(f4.columns), 'Total Attendance Count', '')
        f4.insert(len(f4.columns), 'Real', '')
        f4.insert(len(f4.columns), 'Duplicate', '')
        f4.insert(len(f4.columns), 'Invalid', '')
        f4.insert(len(f4.columns), 'Absent', '')
        index_f4 = 0
        for j in total_lectures:

            #Before writing data, we need to create a blank row otherwise python will raise IndexError
            # pandas.Series() will create a new series 
            # it takes two parameters
            # i) a list of values, all are blank so None
            # ii) index = [Column list to align with columns]
            s = pandas.Series([None,None,None,None,None,None,None,None],index=['Date','Roll','Name','Total Attendance Count','Real','Duplicate','Invalid','Absent'])
            # appending to dataFrame
            f4 = f4.append(s,ignore_index=True)
            #Writing data
            f4['Date'][index_f4] = j
            f4['Roll'][index_f4] = i
            f4['Name'][index_f4] = roll_to_name[i]
            if(j in total_lectures_taken[i].keys()):
                f4['Total Attendance Count'][index_f4] = total_lectures_taken[i][j]
            else :
                f4['Total Attendance Count'][index_f4] = 0
            pass
            #Absent is opposite to presence of real attendance 
            if(j in attendance_real[i]):
                f4['Real'][index_f4] = 1
                f4['Absent'][index_f4] = 0

            else :
                f4['Real'][index_f4] = 0
                f4['Absent'][index_f4] = 1
            pass 
            #Writing duplicate attendance
            if(j in attendance_duplicate[i].keys()):
                f4['Duplicate'][index_f4] = attendance_duplicate[i][j]
            else :
                f4['Duplicate'][index_f4] = 0
            pass
            #writing invalid attendance
            if(j in attendance_invalid[i].keys()):
                f4['Invalid'][index_f4] = attendance_invalid[i][j]
            else :
                f4['Invalid'][index_f4] = 0
            pass
            index_f4 += 1
        


            
       
        #now for consolidated report
        dummy = [None for i in range(5 + len(total_lectures))]
        s = pandas.Series(dummy, index = list(f3.columns)) 
        #Add series
        f3 = f3.append(s, ignore_index = True)
        n_present = 0
        #writing all the values for current roll number in the inserted blank row
        f3['Roll'][index_f3] = i
        f3['Name'][index_f3] = roll_to_name[i]
        #Inserting for dynamic date columns through for loop
        for j in total_lectures:
            if(j in attendance_real[i]):
                f3[j][index_f3] = 'P'
                n_present += 1
            else:
                f3[j][index_f3] = 'A'
        f3['Actual Lecture Taken'][index_f3] = len(total_lectures)
        f3['Total Real'][index_f3] = n_present
        f3['% Attendance'][index_f3] = round(100.00 * n_present / len(total_lectures),2)
        index_f3 += 1
        
        #Saving individual attendance report
        #Since we are saving this file in a subfolder,
        #Method to save in specific folder: os.path.join method, for joining one or more path components.
        #Also since name of file is roll number, so using fstring to genereate name of file
        try:
            f4.to_excel(os.path.join('output',f'{i}.xlsx'), sheet_name=i, index=False)
        except FileExistsError:
            print('File Already exist')
        except: 
            print('error in saving individual report')
    
    #saving consolidated report
    #Method to save in specific folder: os.path.join method, for joining one or more path components.
    try:
        f3.to_excel(os.path.join('output','attendance_report_consolidated.xlsx'), sheet_name="Consolidated Report", index= False)
    except FileExistsError:
        print('File Already Exists')
    except:
        print('error in saving consolidated report')
    # #Freeing up the memory, just a good practice
    del myset, mydict, attendance_real, attendance_duplicate, attendance_invalid, total_lectures, total_lectures_taken

    
    

###Code

from platform import python_version
ver = python_version()

if ver == "3.8.10":
    print("Correct Version Installed")
else:
    print("Please install 3.8.10. Instruction are present in the GitHub Repo/Webmail. Url: https://pastebin.com/nvibxmjw")


attendance_report()
#what a ever is the limit of your sending mails, like gmail has 500.
max_count = 9999999
count=0
file_path = "output\attendance_report_consolidated.xlsx"
try:
        if isEmail(receiver):
            send_mail(FROM_ADDR, FROM_PASSWD, receiver, Subject, Body, file_path)
        print("Count Value: ", count)
        print("Sleeping . .. . ")
        sleep(randint(1,3))
except:
        print("Lets see ")





#This shall be the last lines of the code.
end_time = datetime.now()
print('Duration of Program Execution: {}'.format(end_time - start_time))
