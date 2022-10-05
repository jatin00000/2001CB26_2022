

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
def Value_put(Pointer2, mod_range, counter, Pointer_Transition_Range, title):
    #Describing the parameters
    # Pointer2 is a file pointer
    # mod_range is a fstring containing in the format "<start value>-<end value>"
    # title is used to pass the following for printing 'Overall Count' or 'Mod Transition Count'
    # counter is a reference to print various row and column of matrix
    # Pointer_Transition_Range is a pointer to dictionary for particular transition range

    #1 Inserting the values which are fixed
    Pointer2['Octant ID'][counter]=title
    Pointer2['Octant ID'][counter+1]=mod_range
    Pointer2['1'][counter+1]='To'
    Pointer2[' '][counter+3]='From'
    Pointer2['Octant ID'][counter+2]='Count'
    Pointer2['Octant ID'][counter+3]='+1'
    Pointer2['Octant ID'][counter+4]='-1'
    Pointer2['Octant ID'][counter+5]='+2'
    Pointer2['Octant ID'][counter+6]='-2'
    Pointer2['Octant ID'][counter+7]='+3'
    Pointer2['Octant ID'][counter+8]='-3'
    Pointer2['Octant ID'][counter+9]='+4'
    Pointer2['Octant ID'][counter+10]='-4'
    Pointer2['1'][counter+2]='+1'
    Pointer2['-1'][counter+2]='-1'
    Pointer2['2'][counter+2]='+2'
    Pointer2['-2'][counter+2]='-2'
    Pointer2['3'][counter+2]='+3'
    Pointer2['-3'][counter+2]='-3'
    Pointer2['4'][counter+2]='+4'
    Pointer2['-4'][counter+2]='-4'

    #Space_according_to_Octant is a dictionary which contain position for particular transition value column as they need to come in a particular order
    Space_according_to_Octant ={'1':3, '-1':4,'2':5,'-2':6,'3':7,'-3':8,'4':9, '-4':10} #{octant: position along horizontal direction}
    #Feeding main values in the matrix
    for i in ['1','-1','2','-2','3','-3','4','-4']:
        for j in ['1','-1','2','-2','3','-3','4','-4']:
            Pointer2[i][counter+Space_according_to_Octant[j]]=Pointer_Transition_Range[f'{j}{i}']

def Printer(_hash, Pointer2, index, Count_list):
       #Function to Print Rank for each octant
       #Also keep count number of times Rank 1 has occured for each octant

        _hash = dict(sorted(_hash.items(), key= lambda x:x[1], reverse=True))
        #sorts _hash based on the value of key as applied to each element of the list
        #To sort in decreasing fashion, reverse = True is applied
        #It returns as a list and we need it in dictionary format, so converted

        #This dictionary _Rank stores key as octant and value as its Rank
        _Rank = dict()
        Rank = 1 #Just a variable to assign rank to each octant
        #For each Octant, are sorted in decreasing order, so Octant with maximum count will come first
        # +1:414, -1:688, +2:815   ------After Sorting----->>>> +2:815, -1:688, +1:414
        for i in _hash:
            _Rank[i]=Rank  #assign this octant the following rank
            Rank += 1 #Increase the rank to rank+1 as next element will have less number of count

        #On closely looking, one can see that Rank 1 instead of having all Rank 1 octant name written below it
        #Have value of Rank of +1 octant in it
        #Using them as Column headings instead of Octact Symbol
        List_label = {'1':'Rank 1', '-1':'Rank 2', '2':'Rank 3', '-2':'Rank 4', '3':'Rank 5', '-3':'Rank 6', '4':'Rank 7', '-4':'Rank 8'}
        for i in _Rank:
            #For each range of mod, in that particular row having number as index
            #assigning rank for each octant
            #Column come from Above List_label dictionary
            Pointer2[List_label[i]][index]=_Rank[i]

            #If any octant has Rank 1st
            if(_Rank[i]==1):
                #In the column 'Rank1 Octant ID', put this octant
                Pointer2['Rank1 Octant ID'][index]=i
                #In the column 'Rank1 Octant Name', put its mapping
                Pointer2['Rank1 Octant Name'][index]=octant_name_id_mapping[i]
                #Rank 1st has occured for this octant, increase the count of number of times rank 1st has occured for this octant
                Count_list[i] += 1

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

    try:
        #7 For each mod range of each octant, count number of octant in it
        for counter, rows in Pointer2.iterrows():
            # print(counter)
            for idx in Count_range_wise[str(Pointer2['Octant'][counter])]:
                    if counter>=idx[0] and counter<=idx[1]:
                        #if lies in this interval, increment the count
                        idx[2]+=1
                        #we don't need to look furthur
                        break
    except KeyError:
        print('Key Error in Part 7')
    except:
        print('Other Error in Part 7')

    #8 Adding mod range labels in octant_output.csv file
    for count in range(len(Bounds_mod_range)-1):
        #counter+2 as they will start to fill from 2nd row in Octant ID column
        #will insert the range as an string
        Pointer2['Octant ID'][count+2]=str(Count_range_wise['1'][count][0])+'-'+str(Count_range_wise['1'][count][1])

    #9 Filling count of each octant in each mod range into csv file
    for key, value in Count_range_wise.items(): #iterating through dictionary
        for counter in range(len(Bounds_mod_range)-1): #number of ranges of mod will be constant
            Pointer2[str(key)][counter+2]=int(Count_range_wise[key][counter][2])
            #<file_handler>['<column_for_octant>'][row number] = string of Count_range_wise[key is octant][serial number of mod range][count of values]
    Pointer2.rename(columns = {'Dummy':''}, inplace = True)

    #10 Inserting new columns according to requirement
    Pointer2.insert(len(Pointer2.columns), 'Rank 1', '')
    Pointer2.insert(len(Pointer2.columns), 'Rank 2', '')
    Pointer2.insert(len(Pointer2.columns), 'Rank 3', '')
    Pointer2.insert(len(Pointer2.columns), 'Rank 4', '')
    Pointer2.insert(len(Pointer2.columns), 'Rank 5', '')
    Pointer2.insert(len(Pointer2.columns), 'Rank 6', '')
    Pointer2.insert(len(Pointer2.columns), 'Rank 7', '')
    Pointer2.insert(len(Pointer2.columns), 'Rank 8', '')
    Pointer2.insert(len(Pointer2.columns), 'Rank1 Octant ID', '')
    Pointer2.insert(len(Pointer2.columns), 'Rank1 Octant Name', '')

    #11 This dictionary Count_list contain number of times rank 1st occur for each octant
    Count_list ={'1':0, '-1':0, '2':0, '-2':0, '3':0, '-3':0, '4':0, '-4':0}

    #Calling the function Printer() to print Rank for overall count and passing it dictionary _hash, 
    #1st Row will have Rank for overall count so index = 0
    Printer(_hash, Pointer2, 0, Count_list)

    #Since we don't want to include count of Rank 1 of overall count, we redefine it
    Count_list ={'1':0, '-1':0, '2':0, '-2':0, '3':0, '-3':0, '4':0, '-4':0}

    #mod ranges will start from 2nd row
    #This index variable is used to print rank list for each mod range
    index = 2

    try:
        #12 For each mod range
        for i in range(len(Bounds_mod_range)-1):
            #Calling the Printer() Function
            # Octant ID	      +1	-1	     +2
            # Overall Count	 2610	4603	4855
            # Mod 5000			
            # 0-4999	      414	 688	815        <<<<===== Row number 2
            # 5000-9999	      380	 757    820
            # 10000-14999	  621	1016	599
            # 15000-19999	  366	682	    948

            #Since Function needs a dictionary in octant:count form 
            #Directly accessing the cells for each octant for each mod range and passing them as dictionary through pre-defining as number of octant are 8 only
            #Since ranks will be printed in same row, passing (index + i) 
            #As for 2nd mod range, we will do all the operations in (index+1)th row
            Printer({'1':Pointer2['1'][index+i],
            '-1':Pointer2['-1'][index+i],
            '2':Pointer2['2'][index+i],
            '-2':Pointer2['-2'][index+i],
            '3':Pointer2['3'][index+i],
            '-3':Pointer2['-3'][index+i],
            '4':Pointer2['4'][index+i],
            '-4':Pointer2['-4'][index+i],
            }, Pointer2, index+i, Count_list)
    except ValueError:
        print('Value Error in Part 12')
    except :
        print('Other Error in Part 12')

    #Moving to three row below to the row having count of last mod range, Three row space has been left according to specification
    index += 2 + len(Bounds_mod_range)

    #13 There position will remain unaffected by other things so directly putting values in cells
    Pointer2['1'][index] = 'Octant ID'
    Pointer2['1'][index+1] = '1'
    Pointer2['1'][index+2] = '-1'
    Pointer2['1'][index+3] = '2'
    Pointer2['1'][index+4] = '-2'
    Pointer2['1'][index+5] = '3'
    Pointer2['1'][index+6] = '-3'
    Pointer2['1'][index+7] = '4'
    Pointer2['1'][index+8] = '-4'
    Pointer2['-1'][index] = 'Octant Name'

    #Putting the mapping of octants
    Pointer2['-1'][index+1] = octant_name_id_mapping[ '1']
    Pointer2['-1'][index+2] = octant_name_id_mapping[ '-1']
    Pointer2['-1'][index+3] = octant_name_id_mapping[ '2']
    Pointer2['-1'][index+4] = octant_name_id_mapping[ '-2']
    Pointer2['-1'][index+5] = octant_name_id_mapping[ '3']
    Pointer2['-1'][index+6] = octant_name_id_mapping[ '-3']
    Pointer2['-1'][index+7] = octant_name_id_mapping[ '4']
    Pointer2['-1'][index+8] = octant_name_id_mapping[ '-4']

    #Putting the number of times Rank 1st has occured for each octant
    Pointer2['2'][index] = 'Count of Rank 1 Mod Values'
    Pointer2['2'][index+1] = Count_list[ '1']
    Pointer2['2'][index+2] = Count_list[ '-1']
    Pointer2['2'][index+3] = Count_list[ '2']
    Pointer2['2'][index+4] = Count_list[ '-2']
    Pointer2['2'][index+5] = Count_list[ '3']
    Pointer2['2'][index+6] = Count_list[ '-3']
    Pointer2['2'][index+7] = Count_list[ '4']
    Pointer2['2'][index+8] = Count_list[ '-4']

    #14
    """
    According to Output file, The header should be like
    According to output file:::::::
    <empty cells>	|   |     |     |       |       |       |	1 |	-1 |	2| 	-2| 3|  .... so on
    Time | U |	V |	W	| U Avg |	V Avg |	W Avg |	U'=U - U avg| ...... so on........
      0  |32.8| 5.97|	3.08|	27.30653354 |	0.048330476 |	0.107577744	 | 5.493466465 | ..... so on

    My current output:::::::::::::::::::
    Time | U |	V |	W	| U Avg |	V Avg |	W Avg |	U'=U - U avg| ...... so on........
      0  |32.8| 5.97|	3.08|	27.30653354 |	0.048330476 |	0.107577744	 | 5.493466465 | ..... so on

    So i basically inserted a new row at index 0 containing names of columes as value for each column

    Pandas.concat() will make a new dataframe by
    Time | U |	V |	W	| U Avg |	V Avg |	W Avg |	U'=U - U avg| ...... so on........      <<<<<<<================ This new row
           + Plus
    old dataframe Pointed by Pointer2 variable

    ===> Time | U |	V |	W	| U Avg |	V Avg |	W Avg |	U'=U - U avg| ...... so on........
                +
         Time | U |	V |	W	| U Avg |	V Avg |	W Avg |	U'=U - U avg| ...... so on........
           0  |32.8| 5.97|	3.08|	27.30653354 |	0.048330476 |	0.107577744	 | 5.493466465 | ..... so on
    
    ===> Time | U |	V |	W	| U Avg |	V Avg |	W Avg |	U'=U - U avg| ...... so on........
         Time | U |	V |	W	| U Avg |	V Avg |	W Avg |	U'=U - U avg| ...... so on........
          0  |32.8| 5.97|	3.08|	27.30653354 |	0.048330476 |	0.107577744	 | 5.493466465 | ..... so on
    """

    """
    pnadas.concat(<dataframe 1>, <dataframe 2>)
    <dataframe 1> = list containing columns name and these will be assigned to columns in the sequence respectively
    <file_Pointer>.columns.values will return a list of columns name
    columns = Pointer2.columns will assign each column name from list in same columns itself

    <dataframe 2> is pointing to our old Pointer2 only
    Store the new dataframe in Pointer2 only
    """
    try:
        Pointer2 = pandas.concat([pandas.DataFrame([list(Pointer2.columns.values)],columns=Pointer2.columns),Pointer2],ignore_index=True)
        #Since we have created a new dataframe, so the index of excel needs to be reseted, it is done by below code
        Pointer2 = Pointer2.sort_index().reset_index(drop=True)

        #Now in the header row, all columns name except type 'Rank 1', 'Rank 2, and so on will become null and these will be replaced by '1', '-1', '2' and so on ....
        # Columns names changed using for loop for following names and 
        Pointer2.rename(columns = {x:'' for x in ['Octant ID','Time','U', 'V', 'W', 'U Avg', 'V Avg', 'W Avg','U\'=U - U avg','V\'=V - V avg','W\'=W - W avg', 'Octant', '1','-1','2','-2','3','-3','4','-4','Rank1 Octant ID', 'Rank1 Octant Name' ]}, inplace = True)
        #changing others according to specification 
        Pointer2.rename(columns = {'Rank 1':'1', 'Rank 2':'-1', 'Rank 3':'2', 'Rank 4':'-2', 'Rank 5':'3', 'Rank 6':'-3', 'Rank 7':'4', 'Rank 8':'-4'}, inplace = True)
    except IndexError:
        print('Index Error in Part 14')
    except:
        print('Other Error in Part 14')
    """
    My current look of output file:::::::
    <empty cells>	|   |     |     |       |       |       |	1 |	-1 |	2| 	-2| 3|  .... so on
    Time | U |	V |	W	| U Avg |	V Avg |	W Avg |	U'=U - U avg| ...... so on........
      0  |32.8| 5.97|	3.08|	27.30653354 |	0.048330476 |	0.107577744	 | 5.493466465 | ..... so on
    we are done
    """

    #15 Saving all changes to output file
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
