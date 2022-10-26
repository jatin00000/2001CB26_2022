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
#applied action of 'ignore' to 'FutureWanings' category
import warnings
warnings.simplefilter(action='ignore', category=FutureWarning)

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
    attendance_count_everyday = dict()
    attendance_actual_list, attendance_actual_set = dict(), dict()
    attendance_fake_list, attendance_fake_set  = dict(), dict()
    ''' 
    1. mydict is a simple dictionary which will be used in defining other structures, by making and suppling a copy of it
    2. myset is a set similar to mydict
    3. roll_to_name is a dictionary
        having key : roll number and value : name of this roll number

    4. total_lectures is a set containing valid dates on which lecture happened in form of string

    5. attendance_actual_set is a dictionary of set
        key : Roll number
        value : a set which contains dates stored in form of string for each roll number when a person has actual attendance
        attendance_actual_set = {'2001CB02':{'01/09/2022', '28/07/2022', .....}, '2001CB03':{'01/09/2022', '28/07/2022', .....}, .....}

    6. attendance_fake_set is similar to attendance_actual_set but has dates when attendance of a roll number is considered fake
    7. total_lectures_taken is similar to attendance_actual_set but has dates when attendance is marked on Monday or Thrusday independent of time
    
    Now, attendance_actual_list, attendance_fake_list and attendance_count_everyday are dictionary of dictionary of lists
    key : roll number
    value : another dictionary for which~
        key : date in form of string
        value : a list of Timeframes
    example: 
    {'2001CB02': {'28/07/2022': ['28/07/2022 23:19:31'], '01/08/2022': ['01/08/2022 14:33:36'], .....},
     '2001CB03': {'28/07/2022': ['28/07/2022 17:07:45'], '01/08/2022': ['01/08/2022 14:33:14'], .....},
     .........}

    8. attendance_actual_list has Timeframes on which attendance is counted actual
    9. attendance_fake_list has Timeframes on which attendance is counted as fake
    10. attendance_count_everyday has Timeframes on which Lecture occued on Monday or Thrusday to mark duplicate attendance
    '''

    #Created Regrex Expression for extracting Roll Number, Time, Date in between various values
    roll_number_pattern = re.compile(r'2001[A-Za-z]{2}[\d]{2}')
    time_pattern = re.compile(r'[\d]{2}:[\d]{2}:[\d]{2}')
    date_pattern = re.compile(r'[\d]{2}\/[\d]{2}\/[\d]{4}')


    #2 Initialisation of various data structures
    #iterrows is a function which iterrates all the rows
    #counter or variable at 1st place holds index (0 based indexing)
    #row holds entire row of csv file in form of string
    for counter, rows in f2.iterrows():
        #we can access any row of a particular column by
        #<file_pointer>['<column_label>'][counter] = <value>

        #inserting name in dictionary
        roll_to_name[f2['Roll No'][counter]]=f2['Name'][counter]
        total_lectures_taken[f2['Roll No'][counter]] = myset.copy()
        try:
            #initiallising with default dictionary so that python understand the type of data structure
            attendance_actual_list[f2['Roll No'][counter]] = {" ":[]}
            attendance_actual_set[f2['Roll No'][counter]] = myset.copy()
            attendance_fake_list[f2['Roll No'][counter]] = {" ":[]}
            attendance_fake_set[f2['Roll No'][counter]] = myset.copy()
            attendance_count_everyday[f2['Roll No'][counter]] = {" ": []}
        except KeyError:
            print("KeyError in Part 2")
        except:
            print("Some other Error in Part 2")

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
                
                #Since KeyError was occuring, i adopted this method to add dictionary of lists
                #if the date is not present as key in dictionary, initiallise it with Timestamp
                if(_date not in attendance_count_everyday[roll_number].keys()):
                    attendance_count_everyday[roll_number].setdefault(_date, [Timestamp])
                else :
                    #just append it in 
                    attendance_count_everyday[roll_number][_date].append(Timestamp)
                pass

                #Pandas.Timestamp() is used pass string to pandas, so that it convert it into a Timestamp
                #Weekday() return day of the week on provided Timestamp
                #'0' for Monday and '3' for Thrusday                
                if(pandas.Timestamp(_date[6:]+"-"+_date[3:5]+"-"+_date[0:2]).weekday() in [0,3]):
                    #Adding this date to valid lecture occured
                    total_lectures.add(_date)

                    #Adding to valid lecture count for this roll number
                    total_lectures_taken[roll_number].add(_date)

                    #Extracting hour, minute, second through slicing
                    _Hour, _minute, _second = int(_time[0:2]), int(_time[3:5]), int(_time[6:])

                    #Checking if this attendance is actual or fake by comparing time
                    #Actual is in between 14:00:00 to 15:00:00
                    if( (_Hour==14 and (_minute>=0 and _minute<=59) and (_second>=0 and _second<=59))  or (_Hour == 15 and _minute==0 and _second==0 )):
                        attendance_actual_set[roll_number].add(_date)

                        #Since KeyError was occuring, i adopted this method to add dictionary of lists
                        #if the date is not present as key in dictionary, initiallise it with Timestamp
                        if(_date not in attendance_actual_list[roll_number].keys()):
                            attendance_actual_list[roll_number].setdefault(_date, [Timestamp])
                        else :
                            #just append it in
                            attendance_actual_list[roll_number][_date].append(Timestamp)                    
                    else :
                        #Attendance is fake
                        attendance_fake_set[roll_number].add(_date)

                        #Since KeyError was occuring, i adopted this method to add dictionary of lists
                        #if the date is not present as key in dictionary, initiallise it with Timestamp
                        if(_date not in attendance_fake_list[roll_number].keys()):
                            attendance_fake_list[roll_number].setdefault(_date, [Timestamp])
                        else :
                             #just append it in
                            attendance_fake_list[roll_number][_date].append(Timestamp)
                else :
                    attendance_fake_set[roll_number].add(_date)
                    if(_date not in attendance_fake_list[roll_number].keys()):
                        attendance_fake_list[roll_number].setdefault(_date, [Timestamp])
                    else :
                        attendance_fake_list[roll_number][_date].append(Timestamp)

    #Deleting the default empty key we inputted during initialisation
    for i in roll_to_name.keys():
        del attendance_actual_list[i][" "]
        del attendance_fake_list[i][" "]
        del attendance_count_everyday[i][" "]
    

    
    

###Code

from platform import python_version
ver = python_version()

if ver == "3.8.10":
    print("Correct Version Installed")
else:
    print("Please install 3.8.10. Instruction are present in the GitHub Repo/Webmail. Url: https://pastebin.com/nvibxmjw")


attendance_report()




#This shall be the last lines of the code.
end_time = datetime.now()
print('Duration of Program Execution: {}'.format(end_time - start_time))
