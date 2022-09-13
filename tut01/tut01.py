#Import the necessary modules
#I will be reading file through using csv, pandas both
import csv
import os
from csv import writer, reader
import pandas as pd

#As pandas raise settingWithCopyWarning,
#As it occur due to chain assignment,
#so to avoid its display in terminal, i used options attribute which is used to configure and prevent it from displaying error and exception
pd.options.mode.chained_assignment = None

#Octant function takes three coordinates as parameter and return the octant to which these belong
def octant(x,y,z):
    if z>=0.0:
        if(x>=0.000000000 and y>=0.000000000  ): return 1
        elif(x<0.000000000  and y>=0.000000000 ): return 2
        elif(x<0.000000000  and y<0.000000000 ): return 3
        elif(x>=0.000000000 and y<0.000000000 ): return 4
    else:
        if(x>=0.000000000  and y>=0.000000000 ): return -1
        elif(x<0.000000000  and y>=0.000000000 ): return -2
        elif(x<0.000000000 and y<0.000000000 ): return -3
        elif(x>=0.000000000  and y<0.000000000 ): return -4

def octact_identification(mod=5000):
    f1 = open('octant_input.csv', 'r')
    f2 = open('octant_output.csv', 'a')
    #1counting the number of rows in csv file and computing the average of each U, V, W
    with f1:
        avg_U, avg_V,avg_W=0.0,0.0,0.0 #first store the sum of all observation
        row__=0 #variable to keep count of number of rows
        f2.writelines(f1.readlines()) #writing the first 4 columns in octant_output.csv
        f1.seek(0) #move pointer in file to beginning of file
        for row_number, row in enumerate(f1.readlines()):
            #enumerate return a key: number of row and pair:row in string form
            if row_number == 0:
                continue #ignoring first row, it contains headers
            else:
                row__ += 1
                data = row.split(',')
                avg_U = round(float(data[1]),9)+avg_U
                avg_V = round(float(data[2]),9)+avg_V
                avg_W = round(float(data[3]),9)+avg_W
        #computing average which is total sum of observations / number of observation
        avg_U= round(1.0*avg_U/row__,9) #round to 9 decimal places
        avg_V= round(1.0*avg_V/row__,9)
        avg_W= round(1.0*avg_W/row__,9)
    f1.close(), f2.close() #close both file handlers for now

    #2 Opening file using pandas
    file1 = pd.read_csv("octant_input.csv")
    file2 = pd.read_csv("octant_output.csv")

    #inserting using insert(index to insert, value of column, value by which every cell will be filled)
    #creating seven column and inserting blank value
    file2.insert(len(file2.columns), 'U avg', '')
    file2.insert(len(file2.columns), 'V avg', '')
    file2.insert(len(file2.columns), 'W avg', '')
    file2.insert(len(file2.columns), 'U\'=U - U avg', '')
    file2.insert(len(file2.columns), 'V\'=V - V avg', '')
    file2.insert(len(file2.columns), 'W\'=W - W avg', '')
    file2.insert(len(file2.columns), 'Octant', '')

    #Rounding the column value upto 9 decimal values
    file2['U avg'][0]=round(avg_U,9)
    file2['V avg'][0]=round(avg_V,9)
    file2['W avg'][0]=round(avg_W,9)

    #creating a dictionary to keep count of the eight octant
    _hash = {'1':0, '-1':0, '2':0, '-2':0, '3':0, '-3':0, '4':0, '-4':0}
    
    #iterrows() a similar function as enumerate()
    for index, rows in file2.iterrows():
        #we can access any row of a particular column by
        #<file_pointer>['<column_label>'][index] = <value>
        file2['U\'=U - U avg'][index]=round(float(file2['U'][index]),9)-round(avg_U,9) #subtracting individual reading from average
        file2['V\'=V - V avg'][index]=round(float(file2['V'][index]),9)-round(avg_V,9)
        file2['W\'=W - W avg'][index]=round(float(file2['W'][index]),9)-round(avg_W,9)
        #calling the octant() function to give octant value, and storing it to the cell in octant colunn
        file2['Octant'][index]=octant(round(float(file2['U\'=U - U avg'][index]),9), round(float(file2['V\'=V - V avg'][index]),9), round(float(file2['W\'=W - W avg'][index]),9))
        #increasing the octant count in dictionary
        _hash[str(file2['Octant'][index])] +=1

    #3 Adding the remaining columns same as adding columns U avg, V avg
    file2.insert(len(file2.columns), ' ', '')
    file2.insert(len(file2.columns), 'Octant ID', '')
    file2.insert(len(file2.columns), '1', '')
    file2.insert(len(file2.columns), '-1', '')
    file2.insert(len(file2.columns), '2', '')
    file2.insert(len(file2.columns), '-2', '')
    file2.insert(len(file2.columns), '3', '')
    file2.insert(len(file2.columns), '-3', '')
    file2.insert(len(file2.columns), '4', '')
    file2.insert(len(file2.columns), '-4', '')

    #4 Filling the cells with values 'User Input', 'Overall count' as there position remain fixed whatever be the mod
    file2[' '][1]='User Input'
    file2['Octant ID'][1]='Mod '+str(mod)
    file2['Octant ID'][0]='Overall Count'

    #5 Writing the overall count for each octant at respective position
    #There position will also remain fix, independent of value of mod
    file2['1'][0]=_hash['1']
    file2['-1'][0]=_hash['-1']
    file2['2'][0]=_hash['2']
    file2['-2'][0]=_hash['-2']
    file2['3'][0]=_hash['3']
    file2['-3'][0]=_hash['-3']
    file2['4'][0]=_hash['4']
    file2['-4'][0]=_hash['-4']

    #6 Generating the values that will be bound of intervals like 0,5000,10000 for mod = 5000 and storing them in list
    list_value=[]
    for i in range(0, row__, mod): #start with zero till number of rows and increment by mod
        list_value.append(i)
    else: #at last, the smallest multiple of mod which is greater than number of rows; else will execute at last, able to compute last value of bound
        list_value.append(i+mod)

    #in the upcoming for loop, start position is always incremented by 1
    #so making 0 as -1, -1+1 will again be zero
    list_value.pop(0)
    list_value.insert(0, -1.0000)

    #Again making a dictionary, that will have for each octant all the intervals in form of list
    #Example {'1':[[0,5000,0], [5001,10000,0], .....], '-1':[[0,5000,0], [5001,10000,0], .....], ..........}
    #        'octant_value' : [[start_index, end_index, count of value]]
    list_final = {'1':[], '-1':[], '2':[], '-2':[], '3':[], '-3':[], '4':[], '-4':[]}
    for i in range(len(list_value)-1):
        for j in list_final:
            list_final[j].append([list_value[i]+1,list_value[i+1],0])

    #7 For each mod range of each octant, count number of octant in it
    for index, rows in file2.iterrows():
        for i in list_final[str(file2['Octant'][index])]:
            if index>=i[0] and index<=i[1]:
                #if lies in this interval, increment the count
                i[2]+=1
                #we don't need to look furthur
                break

    
    file2.to_csv("octant_output.csv", index=False)

###Code
from platform import python_version
ver = python_version()


if ver == "3.8.10":
    print("Correct Version Installed")
else:
    print("Please install 3.8.10. Instruction are present in the GitHub Repo/Webmail. Url: https://pastebin.com/nvibxmjw")

#taking user input for mod and calling octact _identification() function
mod = int(input("Enter the mod value:"))
octact_identification(mod)