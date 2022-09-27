#Help https://youtu.be/H37f_x4wAC0
#Import the necessary modules
#I will be reading file through using Pandas
import pandas

#As pandas raise settingWithCopyWarning,
#As it occur due to chain assignment,
#so to avoid its display in terminal, i used options attribute which is used to configure and prevent it from displaying error and exception
pandas.options.mode.chained_assignment = None


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
        Pointer2 = pandas.read_excel("output_octant_longest_subsequence.xlsx")
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