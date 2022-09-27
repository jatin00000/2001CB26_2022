#Help https://youtu.be/H37f_x4wAC0
#Import the necessary modules
#I will be reading file through using Pandas
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
def octant_longest_subsequence_count():
    try:
        Pointer1 = pandas.read_excel("input_octant_longest_subsequence.xlsx")
        Pointer1.to_excel("output_octant_longest_subsequence.xlsx", index=False)
        Pointer2 = pandas.read_excel("output_octant_longest_subsequence.xlsx")
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
        #creating Eleven column and inserting blank value
        Pointer2.insert(len(Pointer2.columns), 'U Avg', '')
        Pointer2.insert(len(Pointer2.columns), 'V Avg', '')
        Pointer2.insert(len(Pointer2.columns), 'W Avg', '')
        Pointer2.insert(len(Pointer2.columns), 'U\'=U - U avg', '')
        Pointer2.insert(len(Pointer2.columns), 'V\'=V - V avg', '')
        Pointer2.insert(len(Pointer2.columns), 'W\'=W - W avg', '')
        Pointer2.insert(len(Pointer2.columns), 'Octant', '')
        Pointer2.insert(len(Pointer2.columns), ' ', '')
        Pointer2.insert(len(Pointer2.columns), 'Count', '')
        Pointer2.insert(len(Pointer2.columns), 'Longest Subsquence Length', '')
        Pointer2.insert(len(Pointer2.columns), 'Count1', '')

        #Rounding the column value upto 9 decimal values
        Pointer2['U Avg'][0]=round(avg_U,9)
        Pointer2['V Avg'][0]=round(avg_V,9)
        Pointer2['W Avg'][0]=round(avg_W,9)
        
        try :
            #iterrows() a similar function as enumerate()
            for counter, rows in Pointer2.iterrows():
                #we can access any row of a particular column by
                #<file_pointer>['<column_label>'][counter] = <value>
                Pointer2['U\'=U - U avg'][counter]=('{:.9f}'.format(float(Pointer2['U'][counter])-avg_U)) #subtracting individual reading from average
                Pointer2['V\'=V - V avg'][counter]=('{:.9f}'.format(float(Pointer2['V'][counter])-avg_V))
                Pointer2['W\'=W - W avg'][counter]=('{:.9f}'.format(float(Pointer2['W'][counter])-avg_W))
                #calling the octant() function to give octant value, and storing it to the cell in octant colunn
                Pointer2['Octant'][counter]=octant(round(float(Pointer2['U\'=U - U avg'][counter]),9), round(float(Pointer2['V\'=V - V avg'][counter]),9), round(float(Pointer2['W\'=W - W avg'][counter]),9))
        except TypeError:
            print('TypeError in Part 2')
        except ValueError:
            print('ValueError in Part 2')
        except :
            print('Some Other type error in Part 2')

        #3Counting longest Subsequence for each octant
        Dict_longes_Sub_seq = { '1':[0,0],'-1':[0,0],'2':[0,0],'-2':[0,0],'3':[0,0],'-3':[0,0],'4':[0,0],'-4':[0,0]}
        # here key = Octant Value and Value  = a list which 0th index is length of longest subsequence for this octant and 
        # number of times this longest sequence has occured at 1st index

        try:
            #starting from first index
            index = 0
            #It should always be two less than number of rows, one due to header and other as indexing starts from zero
            while index<(row__-1):
                #Storing the octant for which, we will find the current longest subsequence
                cur = str(Pointer2['Octant'][index])
                #starts with length zero initially
                length = 0
                #keep running this while loop until we find character same as our cur octant
                while str(Pointer2['Octant'][index]) == cur and index<(row__-1):
                    length += 1 #Increase current subsequence length
                    index += 1 #Move to next index
                
                #If our current subsequence length is greater than previous Greatest Subsequence for current octant
                if length > Dict_longes_Sub_seq[cur][0]:
                    Dict_longes_Sub_seq[cur][0] = length #Make current subsequence length as Greatest Subsequence for current octant
                    Dict_longes_Sub_seq[cur][1] = 1 #It has occured first time, so count is 1
            
                #Else If our current subsequence length is equal to previous Greatest Subsequence for current octant
                elif length == Dict_longes_Sub_seq[cur][0]:
                    #Just Increment the count for this length of Subsequence
                    Dict_longes_Sub_seq[cur][1] += 1
        except ValueError:
            Print('Value Error in Part 3')
        except :
            Print('Some Other Errot=r in Part 3')

        

        #4 Saving all changes to output file
        Pointer2.to_excel("output_octant_longest_subsequence.xlsx", index=False)
    except FileNotFoundError():
        print("File Not Found")
    except: 
        Print("Error Opening the file")
###Code

from platform import python_version
ver = python_version()

if ver == "3.8.10":
    print("Correct Version Installed")
else:
    print("Please install 3.8.10. Instruction are present in the GitHub Repo/Webmail. Url: https://pastebin.com/nvibxmjw")


octant_longest_subsequence_count()