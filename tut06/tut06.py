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
