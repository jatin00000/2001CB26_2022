import csv
import os
from csv import writer
from csv import reader
import pandas as pd

# os.system('cls')
def octant(x,y,z):
    if z>=0.0:
        if(x>=0.0 and y>=0.0): return 1
        elif(x<0.0 and y>=0.0): return 2
        elif(x<0.0 and y<0.0): return 3
        elif(x>=0.0 and y<0.0): return 4
    else:
        if(x>=0.0 and y>=0.0): return -1
        elif(x<0.0 and y>=0.0): return -2
        elif(x<0.0 and y<0.0): return -3
        elif(x>=0.0 and y<0.0): return -4
def octact_identification(mod=5000):

    f1 = open('octant_input.csv', 'r')

    f2 = open('octant_output.csv', 'a')

    #1counting the number of rows in csv file

    with f1:

        temp1, temp2,temp3=0.0,0.0,0.0

        row__=0

        f2.writelines(f1.readlines())
        f1.seek(0)
        for row_number, row in enumerate(f1.readlines()):
            

            if row_number == 0:

                continue

            else:

                row__ += 1

                data = row.split(',')

                temp1 = float(data[1])+temp1

                temp2 = float(data[2])+temp2

                temp3 = float(data[3])+temp3

        data = list(f1.readlines())

        temp1= temp1/row__

        temp2= temp2/row__

        temp3= temp3/row__

        print(temp1,temp2,temp3)
    f1.close(), f2.close()
    file1 = pd.read_csv("octant_input.csv")
    file2 = pd.read_csv("octant_output.csv")
    file2.insert(len(file2.columns), 'U avg', '')
    file2.insert(len(file2.columns), 'V avg', '')
    file2.insert(len(file2.columns), 'W avg', '')
    file2['U avg'][0]=temp1
    file2['V avg'][0]=temp2
    file2['W avg'][0]=temp3
    file2.insert(len(file2.columns), 'U\'=U - U avg', '')
    file2.insert(len(file2.columns), 'V\'=V - V avg', '')
    file2.insert(len(file2.columns), 'W\'=W - W avg', '')
    file2.insert(len(file2.columns), 'Octant', '')
    for index, rows in file2.iterrows():
        file2['U\'=U - U avg'][index]=float(file2['U'][index])-temp1
        file2['V\'=V - V avg'][index]=float(file2['V'][index])-temp2
        file2['W\'=W - W avg'][index]=float(file2['W'][index])-temp3
        file2['Octant'][index]=octant(float(file2['U\'=U - U avg'][index]), float(file2['V\'=V - V avg'][index]), float(file2['W\'=W - W avg'][index]))
    file2.to_csv("octant_output.csv", index=False)

        


    


###Code



from platform import python_version

ver = python_version()


if ver == "3.8.10":

    print("Correct Version Installed")

else:

    print("Please install 3.8.10. Instruction are present in the GitHub Repo/Webmail. Url: https://pastebin.com/nvibxmjw")


mod=5000
octact_identification(mod)