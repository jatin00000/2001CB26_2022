#Import the necessary modules
#I will be reading file through using csv, pandas both
import pandas

#As pandas raise settingWithCopyWarning,
#As it occur due to chain assignment,
#so to avoid its display in terminal, i used options attribute which is used to configure and prevent it from displaying error and exception
pandas.options.mode.chained_assignment = None

#Octant function takes three coordinates as parameter and return the octant to which these belong
def octant(x,y,z):
    if z>=0.000000000:
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
    f1 = pandas.read_csv("octant_input.csv")
    f1.to_csv("octant_output.csv", index=False)
    file2 = pandas.read_csv("octant_output.csv")
    #1counting the number of rows in csv file and computing the average of each U, V, W

    avg_U, avg_V,avg_W=0.0,0.0,0.0 #first store the sum of all observation
    #shape attribute of pandas return number of rows and columns
    row__, column__ = file2.shape #variable to keep count of number of rows

    avg_U = file2['U'].sum()
    avg_V = file2['V'].sum()
    avg_W = file2['W'].sum()
    #computing average which is total sum of observations / number of observation
    avg_U= round(1.0*avg_U/row__,9) #round to 9 decimal places
    avg_V= round(1.0*avg_V/row__,9)
    avg_W= round(1.0*avg_W/row__,9)

    #2inserting using insert(position to insert, value of column, value by which every cell will be filled)
    #creating seven column and inserting blank value
    file2.insert(len(file2.columns), 'U Avg', '')
    file2.insert(len(file2.columns), 'V Avg', '')
    file2.insert(len(file2.columns), 'W Avg', '')
    file2.insert(len(file2.columns), 'U\'=U - U avg', '')
    file2.insert(len(file2.columns), 'V\'=V - V avg', '')
    file2.insert(len(file2.columns), 'W\'=W - W avg', '')
    file2.insert(len(file2.columns), 'Octant', '')

    #Rounding the column value upto 9 decimal values
    file2['U Avg'][0]=round(avg_U,9)
    file2['V Avg'][0]=round(avg_V,9)
    file2['W Avg'][0]=round(avg_W,9)

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
    Bounds_mod_range=[]
    for count in range(0, row__, mod): #start with zero till number of rows and increment by mod
        Bounds_mod_range.append(count)
    Bounds_mod_range.append(row__+1)

    #Again making a dictionary, that will have for each octant all the intervals in form of list
    #Example {'1':[[0,4999,0], [5000,9999,0], .....], '-1':[[0,4999,0], [5000,9999,0], .....], ..........}
    #        'octant_value' : [[start_index, end_index, count of value]]
    Count_range_wise = {'1':[], '-1':[], '2':[], '-2':[], '3':[], '-3':[], '4':[], '-4':[]}
    for index in range(len(Bounds_mod_range)-1):
        for count in Count_range_wise:
            Count_range_wise[count].append([Bounds_mod_range[index],Bounds_mod_range[index+1]-1,0])

    #7 For each mod range of each octant, count number of octant in it
    for index, rows in file2.iterrows():
        # print(index)
        for idx in Count_range_wise[str(file2['Octant'][index])]:
                if index>=idx[0] and index<idx[1]:
                    #if lies in this interval, increment the count
                    idx[2]+=1
                    #we don't need to look furthur
                    break

    #8 Adding mod range labels in octant_output.csv file
    for count in range(len(Bounds_mod_range)-1):
        #index+2 as they will start to fill from 2nd row in Octant ID column
        #will insert the range as an string
        file2['Octant ID'][count+2]=str(Count_range_wise['1'][count][0])+'-'+str(Count_range_wise['1'][count][1])

    #9 Filling count of each octant in each mod range into csv file
    for key, value in Count_range_wise.items(): #iterating through dictionary
        for index in range(len(Bounds_mod_range)-1): #number of ranges of mod will be constant
            file2[str(key)][index+2]=str(Count_range_wise[key][index][2])
            #<file_handler>['<column_for_octant>'][row number] = string of Count_range_wise[key is octant][serial number of mod range][count of values]
    #10 Saving all changes to output file
    file2.to_csv("octant_output.csv", index=False)

###Code
from platform import python_version
ver = python_version()


if ver == "3.8.10":
    print("Correct Version Installed")
else:
    print("Please install 3.8.10. Instruction are present in the GitHub Repo/Webmail. Url: https://pastebin.com/nvibxmjw")

#taking user input for mod and calling octact _identification() function
mod = 5000
octact_identification(mod)