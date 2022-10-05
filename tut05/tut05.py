

from datetime import datetime
start_time = datetime.now()

#Help https://youtu.be/N6PBd4XdnEw
#Import the necessary modules
import pandas

#This line of Pandas->inputoutput->format->file_type = Excel-> ExcelFormatter is used to format excel way of looking
#Header_Style determines the header row look and all the layout related to header
#Setting it none make header row look likes normal data row only
#It is done on specific requirement of appearance of Output, as specified by Sir
pandas.io.formats.excel.ExcelFormatter.header_style = None


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

octant_name_id_mapping = {"1":"Internal outward interaction", "-1":"External outward interaction", "2":"External Ejection", "-2":"Internal Ejection", "3":"External inward interaction", "-3":"Internal inward interaction", "4":"Internal sweep", "-4":"External sweep"}
def octant_range_names(mod=5000):
    Pointer1 = pandas.read_excel("octant_input.xlsx")
    Pointer1.to_excel("octant_output_ranking_excel.xlsx", sheet_name='octant_output_Ranking' ,index = False)
    Pointer2 = pandas.read_excel("octant_output_ranking_excel.xlsx", sheet_name='octant_output_Ranking')
    #1counting the number of rows in csv file and computing the average of each U, V, W
    #shape attribute of pandas return number of rows and columns
    row__, column__ = Pointer2.shape #variable to keep count of number of rows

    avg_U = Pointer2['U'].sum()
    avg_V = Pointer2['V'].sum()
    avg_W = Pointer2['W'].sum()
    #computing average which is total sum of observations / number of observation
    avg_U= round(1.0*avg_U/row__,9) #round to 9 decimal places
    avg_V= round(1.0*avg_V/row__,9)
    avg_W= round(1.0*avg_W/row__,9)

    #2inserting using insert(position to insert, value of column, value by which every cell will be filled)
    #creating seven column and inserting blank value
    Pointer2.insert(len(Pointer2.columns), 'U Avg', '')
    Pointer2.insert(len(Pointer2.columns), 'V Avg', '')
    Pointer2.insert(len(Pointer2.columns), 'W Avg', '')
    Pointer2.insert(len(Pointer2.columns), 'U\'=U - U avg', '')
    Pointer2.insert(len(Pointer2.columns), 'V\'=V - V avg', '')
    Pointer2.insert(len(Pointer2.columns), 'W\'=W - W avg', '')
    Pointer2.insert(len(Pointer2.columns), 'Octant', '')

    #Rounding the column value upto 9 decimal values
    Pointer2['U Avg'][0]=round(avg_U,9)
    Pointer2['V Avg'][0]=round(avg_V,9)
    Pointer2['W Avg'][0]=round(avg_W,9)

    #creating a dictionary to keep count of the eight octant
    _hash = {'1':0, '-1':0, '2':0, '-2':0, '3':0, '-3':0, '4':0, '-4':0}
    
    try:
        #iterrows() a similar function as enumerate()
        for counter, rows in Pointer2.iterrows():
            #we can access any row of a particular column by
            #<file_pointer>['<column_label>'][counter] = <value>
            Pointer2['U\'=U - U avg'][counter]=('{:.9f}'.format(float(Pointer2['U'][counter])-avg_U)) #subtracting individual reading from average
            Pointer2['V\'=V - V avg'][counter]=('{:.9f}'.format(float(Pointer2['V'][counter])-avg_V))
            Pointer2['W\'=W - W avg'][counter]=('{:.9f}'.format(float(Pointer2['W'][counter])-avg_W))
            #calling the octant() function to give octant value, and storing it to the cell in octant colunn
            Pointer2['Octant'][counter]=octant(round(float(Pointer2['U\'=U - U avg'][counter]),9), round(float(Pointer2['V\'=V - V avg'][counter]),9), round(float(Pointer2['W\'=W - W avg'][counter]),9))
            #increasing the octant count in dictionary
            _hash[str(Pointer2['Octant'][counter])] +=1
    except ValueError:
        print('Value error in Part 2')
    except :
        print('Other error in Part 2')

    #3 Adding the remaining columns same as adding columns U avg, V avg
    Pointer2.insert(len(Pointer2.columns), 'Dummy', '')
    Pointer2.insert(len(Pointer2.columns), 'Octant ID', '')
    Pointer2.insert(len(Pointer2.columns), '1', '')
    Pointer2.insert(len(Pointer2.columns), '-1', '')
    Pointer2.insert(len(Pointer2.columns), '2', '')
    Pointer2.insert(len(Pointer2.columns), '-2', '')
    Pointer2.insert(len(Pointer2.columns), '3', '')
    Pointer2.insert(len(Pointer2.columns), '-3', '')
    Pointer2.insert(len(Pointer2.columns), '4', '')
    Pointer2.insert(len(Pointer2.columns), '-4', '')

    #4 Filling the cells with values 'User Input', 'Overall count' as there position remain fixed whatever be the mod
    Pointer2['Dummy'][1]='User Input'
    Pointer2['Octant ID'][1]='Mod '+str(mod)
    Pointer2['Octant ID'][0]='Overall Count'

    #5 Writing the overall count for each octant at respective position
    #There position will also remain fix, independent of value of mod
    Pointer2['1'][0]=_hash['1']
    Pointer2['-1'][0]=_hash['-1']
    Pointer2['2'][0]=_hash['2']
    Pointer2['-2'][0]=_hash['-2']
    Pointer2['3'][0]=_hash['3']
    Pointer2['-3'][0]=_hash['-3']
    Pointer2['4'][0]=_hash['4']
    Pointer2['-4'][0]=_hash['-4']

    try:
        #6 Generating the values that will be bound of intervals like 0,5000,10000 for mod = 5000 and storing them in list
        Bounds_mod_range=[]
        for count in range(0, row__, mod): #start with zero till number of rows and increment by mod
            Bounds_mod_range.append(count)
        Bounds_mod_range.append(row__+1)

        #Again making a dictionary, that will have for each octant all the intervals in form of list
        #Example {'1':[[0,4999,0], [5000,9999,0], .....], '-1':[[0,4999,0], [5000,9999,0], .....], ..........}
        #        'octant_value' : [[start_counter, end_counter, count of value]]
        Count_range_wise = {'1':[], '-1':[], '2':[], '-2':[], '3':[], '-3':[], '4':[], '-4':[]}
        for counter in range(len(Bounds_mod_range)-1):
            for count in Count_range_wise:
                Count_range_wise[count].append([Bounds_mod_range[counter],Bounds_mod_range[counter+1]-1,0])
    except TypeError:
        print('TypeError in Part 6')
    except ValueError:
        print('Value Error in Part 6')
    except:
        print('Other error in Part 6')

    
    #7 Saving all changes to output file
    Pointer2.to_excel("octant_output_ranking_excel.xlsx", sheet_name='octant_output_Ranking' ,index = False)
    
    

###Code

from platform import python_version
ver = python_version()

if ver == "3.8.10":
    print("Correct Version Installed")
else:
    print("Please install 3.8.10. Instruction are present in the GitHub Repo/Webmail. Url: https://pastebin.com/nvibxmjw")


mod=5000 
octant_range_names(mod)



#This shall be the last lines of the code.
end_time = datetime.now()
print('Duration of Program Execution: {}'.format(end_time - start_time))
