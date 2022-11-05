from datetime import datetime
start_time = datetime.now()

#Help
#Help https://youtu.be/N6PBd4XdnEw
#Import the necessary modules and functions from them
import pandas, openpyxl, glob, os, numpy
from openpyxl.styles.borders import Border, Side
from openpyxl.styles import PatternFill
#This line of Pandas->inputoutput->format->file_type = Excel-> ExcelFormatter is used to format excel way of looking
#Header_Style determines the header row look and all the layout related to header
#Setting it none make header row look likes normal data row only
#It is done on specific requirement of appearance of Output, as specified by Sir
pandas.io.formats.excel.ExcelFormatter.header_style = None


#As pandas raise settingWithCopyWarning,
#As it occur due to chain assignment,
#so to avoid its display in terminal, i used options attribute which is used to configure and prevent it from displaying error and exception
pandas.options.mode.chained_assignment = None

#defining border style to apply to the cells as per specification
Myborder = Border(left=Side(style='thin',color='00000000'), 
					  right=Side(style='thin',color='00000000'), 
					  top=Side(style='thin',color='00000000'), 
					  bottom=Side(style='thin',color='00000000'))

#Octant function takes three coordinates as parameter and return the octant to which these belong
def octant(x,y,z):
	if z>=0.000:
		if(x>=0.000 and y>=0.000  ): return 1
		elif(x<0.000  and y>=0.000 ): return 2
		elif(x<0.000  and y<0.000 ): return 3
		elif(x>=0.000 and y<0.000 ): return 4
	else:
		if(x>=0.000  and y>=0.000 ): return -1
		elif(x<0.000  and y>=0.000 ): return -2
		elif(x<0.000 and y<0.000 ): return -3
		elif(x>=0.000  and y<0.000 ): return -4

#This is a dictionary, having column name: column number according to excel indexing which starts from 1
#This will be used later in openpyxl module
col = {'T':1, 'U':2, 'V':3, 'W':4, 'U Avg':5, 'V Avg':6, 'W Avg':7, 'U\'=U - U avg':8, 'V\'=V - V avg':9,'W\'=W - W avg':10, 'Octant':11,'Dummy1':12, 'Dummy':13,'Octant ID':14,'1':15,'-1':16,'2':17,'-2':18,'3':19,'-3':20, '4':21,'-4':22, 'Rank Octant 1':23, 'Rank Octant -1':24,'Rank Octant 2':25,'Rank Octant -2':26, 'Rank Octant 3':27, 'Rank Octant -3':28, 'Rank Octant 4':29, 'Rank Octant -4':30, 'Rank1 Octant ID':31, 'Rank1 Octant Name':32, 'c1':33, 'c2':34, 'c3':35, '1c':36,'-1c':37, '2c':38, '-2c':39, '3c':40, '-3c':41, '4c':42, '-4c':43, 'd1':44, 'd2':45, 'd3':46, 'd4':47, 'e1':48, 'e2':49, 'e3':50, 'e4':51 }



'''
Basically, in my code, all the part that involve writing of data is done using pandas module.
Used openpyxl to add background color and border to cells as per specification.

In openpyxl, cells can be accessed through <worksheet>.cell(row = <row number>, column = <column number>)
and here <row number> and <column number> are integers greater than zero

I had stored <row number> and <column number> in separate lists for cells that will be added border and color.
_border list for cells with border and _color list for cells with color
_border = [[1,2], [10,12], ......] Like this
Later added border and color to them by again opening excel file in openpyxl
'''
def Myadd(_list, _row, _col):
	#Here row number needs to be increased by two because of 0-based indexing and 1st row for header
	#Putting corresponding column number for each column label
	_list.append([int(_row)+2, col[_col]])
	#Modifiying the list and returning back
	return _list
	#Example of calling the function
	#_list = Myadd(_list,2 ,'c3')


octant_name_id_mapping = {"1":"Internal outward interaction", "-1":"External outward interaction", "2":"External Ejection", "-2":"Internal Ejection", "3":"External inward interaction", "-3":"Internal inward interaction", "4":"Internal sweep", "-4":"External sweep"}
def Value_put(Pointer2, mod_range, counter, Pointer_Transition_Range, title, _list, _list2):
	#Describing the parameters
	# Pointer2 is a file pointer
	# mod_range is a fstring containing in the format "<start value>-<end value>"
	# title is used to pass the following for printing 'Overall Count' or 'Mod Transition Count'
	# counter is a reference to print various row and column of matrix
	# Pointer_Transition_Range is a pointer to dictionary for particular transition range
	#_list pointes to _border list
	#_list2 points to _color list

	#1 Inserting the values which are fixed and updating the list for cells that will be added border
	Pointer2['c3'][counter]= title
	Pointer2['c3'][(counter+1)] = mod_range
	Pointer2['1c'][(counter+1)] ='To'
	Pointer2['c2'][counter+3] ='From'
	Pointer2['c3'][counter+2]='Octant #'
	_list = Myadd(_list,counter+2 ,'c3')
	Pointer2['c3'][counter+3]='+1'
	_list = Myadd(_list,counter+3 ,'c3')
	Pointer2['c3'][counter+4]='-1'
	_list = Myadd(_list,counter+4 ,'c3')
	Pointer2['c3'][counter+5]='+2'
	_list = Myadd(_list, counter+5,'c3')
	Pointer2['c3'][counter+6]='-2'
	_list = Myadd(_list,counter+6 ,'c3')
	Pointer2['c3'][counter+7]='+3'
	_list = Myadd(_list, counter+7,'c3')
	Pointer2['c3'][counter+8]='-3'
	_list = Myadd(_list,counter+8 ,'c3')
	Pointer2['c3'][counter+9]='+4'
	_list = Myadd(_list, counter+9,'c3')
	Pointer2['c3'][counter+10]='-4'
	_list = Myadd(_list,counter+10 ,'c3')
	Pointer2['1c'][counter+2]='+1'
	_list = Myadd(_list,counter+2 ,'1c')
	Pointer2['-1c'][counter+2]='-1'
	_list = Myadd(_list,counter+2 ,'-1c')
	Pointer2['2c'][counter+2]='+2'
	_list = Myadd(_list,counter+2 ,'2c')
	Pointer2['-2c'][counter+2]='-2'
	_list = Myadd(_list, counter+2,'-2c')
	Pointer2['3c'][counter+2]='+3'
	_list = Myadd(_list,counter+2 ,'3c')
	Pointer2['-3c'][counter+2]='-3'
	_list = Myadd(_list,counter+2 ,'-3c')
	Pointer2['4c'][counter+2]='+4'
	_list = Myadd(_list,counter+2 ,'4c')
	Pointer2['-4c'][counter+2]='-4'
	_list = Myadd(_list, counter+2,'-4c')

	#Space_according_to_Octant is a dictionary which contain position for particular transition value column as they need to come in a particular order
	Space_according_to_Octant ={'1':3, '-1':4,'2':5,'-2':6,'3':7,'-3':8,'4':9, '-4':10} #{octant: position along horizontal direction}
	#Feeding main values in the matrix
	for i in ['1','-1','2','-2','3','-3','4','-4']:
		#To highlight the max value row wise
		#First storing all the values in a list

		#Finding the max value
		_max = list()
		for j in ['1','-1','2','-2','3','-3','4','-4']:
			_max.append(Pointer_Transition_Range[f'{j}{i}'])
		_Great = max(_max)
		for j in ['1','-1','2','-2','3','-3','4','-4']:
			Pointer2[i+'c'][counter+Space_according_to_Octant[j]] =Pointer_Transition_Range[f'{j}{i}']
			_list = Myadd(_list,counter+Space_according_to_Octant[j],i+'c' )
			#If current value is max, add this cell for highlighting
			if(Pointer_Transition_Range[f'{j}{i}'] == _Great):
				_list2 = Myadd(_list2,counter+Space_according_to_Octant[j] , i+'c')
	return _list, _list2
def Printer(_hash, Pointer2, index, Count_list, _list, _list2):
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
	List_label = {'1':'Rank Octant 1', '-1':'Rank Octant -1', '2':'Rank Octant 2', '-2':'Rank Octant -2', '3':'Rank Octant 3', '-3':'Rank Octant -3', '4':'Rank Octant 4', '-4':'Rank Octant -4'}
	for i in _Rank:
			#For each range of mod, in that particular row having number as index
			#assigning rank for each octant
			#Column come from Above List_label dictionary
		Pointer2[List_label[i]][index] =_Rank[i]
		_list = Myadd(_list, index+1, List_label[i])
		#If any octant has Rank 1st
		if(_Rank[i]==1):
			#Add this cell for highlighting
			_list2 = Myadd(_list2, index+1, List_label[i])
				#In the column 'Rank1 Octant ID', put this octant
			Pointer2['Rank1 Octant ID'][index]=i
			_list = Myadd(_list, index+1, 'Rank1 Octant ID')
				#In the column 'Rank1 Octant Name', put its mapping
			Pointer2['Rank1 Octant Name'][index] = octant_name_id_mapping[i]
			_list = Myadd(_list,index+1, 'Rank1 Octant Name')
				#Rank 1st has occured for this octant, increase the count of number of times rank 1st has occured for this octant
			Count_list[i] += 1
	return _list, _list2



def octant_analysis(mod=5000):
	#used to return all file paths that match a specific pattern. 
	file_list = glob.glob('./input/*.xlsx')
	#Taking each file, one by one
	for myFile in file_list:
		#Creating list for storing cells address for draw border and highlighting
		_border,_color = list(), list()

		#myFile = /input/1.0.xlsx
		#doing splitting and taking the last part i.e. the name
		a = myFile.split("\\")[-1]
		#a = 1.0.xlsx
		Pointer2 = pandas.read_excel(myFile)
		
		#1counting the number of rows in excel file and computing the average of each U, V, W
		#shape attribute of pandas return number of rows and columns
		row__, column__ = Pointer2.shape #variable to keep count of number of rows


		U,V,W =0.0,0.0,0.0
		for counter, rows in Pointer2.iterrows():
			U += float(Pointer2['U'][counter])
			V += float(Pointer2['V'][counter])
			W += float(Pointer2['W'][counter])

		#computing average which is total sum of observations / number of observation
		avg_U= 1.0*U/row__ #round to 9 decimal places
		avg_V= 1.0*V/row__
		avg_W= 1.0*W/row__



		#2inserting using insert(position to insert, value of column, value by which every cell will be filled)
		#creating seven column and inserting blank value
		Pointer2.insert(len(Pointer2.columns), 'U Avg', '')
		Pointer2.insert(len(Pointer2.columns), 'V Avg', '')
		Pointer2.insert(len(Pointer2.columns), 'W Avg', '')
		Pointer2.insert(len(Pointer2.columns), 'U\'=U - U avg', '')
		Pointer2.insert(len(Pointer2.columns), 'V\'=V - V avg', '')
		Pointer2.insert(len(Pointer2.columns), 'W\'=W - W avg', '')
		Pointer2.insert(len(Pointer2.columns), 'Octant', '')



		#creating a dictionary to keep count of the eight octant
		_hash = {'1':0, '-1':0, '2':0, '-2':0, '3':0, '-3':0, '4':0, '-4':0}
		
		try:
			#iterrows() a similar function as enumerate()
			for counter, rows in Pointer2.iterrows():
				#we can access any row of a particular column by
				#<file_pointer>['<column_label>'][counter] = <value>
				Pointer2['U\'=U - U avg'][counter]=('{:.3f}'.format(float(Pointer2['U'][counter])-avg_U)) #subtracting individual reading from average
				Pointer2['V\'=V - V avg'][counter]=('{:.3f}'.format(float(Pointer2['V'][counter])-avg_V))
				Pointer2['W\'=W - W avg'][counter]=('{:.3f}'.format(float(Pointer2['W'][counter])-avg_W))
				#calling the octant() function to give octant value, and storing it to the cell in octant colunn
				Pointer2['Octant'][counter]=octant(round(float(Pointer2['U\'=U - U avg'][counter]),9), round(float(Pointer2['V\'=V - V avg'][counter]),9), round(float(Pointer2['W\'=W - W avg'][counter]),9))
				#increasing the octant count in dictionary
				_hash[str(Pointer2['Octant'][counter])] +=1
		except ValueError:
			print('Value error in Part 2')
		except :
			print('Other error in Part 2')

		#Rounding values in these columns to 3 decimal places
		Pointer2['U']= Pointer2['U'].apply(lambda x: '{:.3f}'.format(x)) 
		Pointer2['V']= Pointer2['V'].apply(lambda x: '{:.3f}'.format(x))
		Pointer2['W']= Pointer2['W'].apply(lambda x: '{:.3f}'.format(x))
		# Pointer2['U\'=U - U avg']= Pointer2['U\'=U - U avg'].apply(lambda x: '{:.3f}'.format(x)) 
		# Pointer2['V\'=V - V avg']= Pointer2['V\'=V - V avg'].apply(lambda x: '{:.3f}'.format(x))
		# Pointer2['W\'=W - W avg']= Pointer2['W\'=W - W avg'].apply(lambda x: '{:.3f}'.format(x))
		#3 Adding the remaining columns same as adding columns U avg, V avg
		Pointer2.insert(len(Pointer2.columns), 'Dummy1', '')
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
		Pointer2.insert(len(Pointer2.columns), 'Rank Octant 1', '')
		Pointer2.insert(len(Pointer2.columns), 'Rank Octant -1', '')
		Pointer2.insert(len(Pointer2.columns), 'Rank Octant 2', '')
		Pointer2.insert(len(Pointer2.columns), 'Rank Octant -2', '')
		Pointer2.insert(len(Pointer2.columns), 'Rank Octant 3', '')
		Pointer2.insert(len(Pointer2.columns), 'Rank Octant -3', '')
		Pointer2.insert(len(Pointer2.columns), 'Rank Octant 4', '')
		Pointer2.insert(len(Pointer2.columns), 'Rank Octant -4', '')
		Pointer2.insert(len(Pointer2.columns), 'Rank1 Octant ID', '')
		Pointer2.insert(len(Pointer2.columns), 'Rank1 Octant Name', '')
		

		#4 Filling the cells with values  'Overall count' as there position remain fixed whatever be the mod
		Pointer2['Dummy'][1]='Mod '+str(mod)
		Pointer2['Octant ID'][0]='Octant ID'
		_border = Myadd(_border, 1,'Octant ID' )
		Pointer2['Octant ID'][1] = 'Overall Count'
		_border = Myadd(_border, 2,'Octant ID' )
		Pointer2['1'][0]='1'
		_border = Myadd(_border, 1, '1' )
		Pointer2['-1'][0]='-1'
		_border = Myadd(_border, 1,'-1' )
		Pointer2['2'][0]='2'
		_border = Myadd(_border, 1, '2')
		Pointer2['-2'][0]='-2'
		_border = Myadd(_border,1 , '-2')
		Pointer2['3'][0]='3'
		_border = Myadd(_border, 1,'3' )
		Pointer2['-3'][0]='-3'
		_border = Myadd(_border, 1,'-3' )
		Pointer2['4'][0]='4'
		_border = Myadd(_border, 1,'4' )
		Pointer2['-4'][0]='-4'
		_border = Myadd(_border, 1,'-4' )
		Pointer2['Rank Octant 1'][0]='Rank Octant 1'
		_border = Myadd(_border, 1,'Rank Octant 1' )
		Pointer2['Rank Octant -1'][0]='Rank Octant -1'
		_border = Myadd(_border, 1,'Rank Octant -1' )
		Pointer2['Rank Octant 2'][0]='Rank Octant 2'
		_border = Myadd(_border, 1,'Rank Octant 2' )
		Pointer2['Rank Octant -2'][0]='Rank Octant -2'
		_border = Myadd(_border, 1,'Rank Octant -2' )
		Pointer2['Rank Octant 3'][0]='Rank Octant 3'
		_border = Myadd(_border, 1,'Rank Octant 3' )
		Pointer2['Rank Octant -3'][0]='Rank Octant -3'
		_border = Myadd(_border, 1,'Rank Octant -3' )
		Pointer2['Rank Octant 4'][0]='Rank Octant 4'
		_border = Myadd(_border, 1,'Rank Octant 4' )
		Pointer2['Rank Octant -4'][0]='Rank Octant -4'
		_border = Myadd(_border, 1,'Rank Octant -4' )
		Pointer2['Rank1 Octant ID'][0]='Rank1 Octant ID'
		_border = Myadd(_border, 1,'Rank1 Octant ID' )
		Pointer2['Rank1 Octant Name'][0]='Rank1 Octant Name'
		_border = Myadd(_border, 1,'Rank1 Octant Name' )

		#5 Writing the overall count for each octant at respective position
		#There position will also remain fix, independent of value of mod
		Pointer2['1'][1]=_hash['1']
		_border = Myadd(_border, 2,'1' )
		Pointer2['-1'][1]=_hash['-1']
		_border = Myadd(_border, 2, '-1')
		Pointer2['2'][1]=_hash['2']
		_border = Myadd(_border, 2, '2')
		Pointer2['-2'][1]=_hash['-2']
		_border = Myadd(_border, 2, '-2')
		Pointer2['3'][1]=_hash['3']
		_border = Myadd(_border, 2, '3')
		Pointer2['-3'][1]=_hash['-3']
		_border = Myadd(_border, 2, '-3' )
		Pointer2['4'][1]=_hash['4']
		_border = Myadd(_border, 2, '4')
		Pointer2['-4'][1]=_hash['-4']
		_border = Myadd(_border, 2, '-4' )



##Read all the excel files in a batch format from the input/ folder. Only xlsx to be allowed
##Save all the excel files in a the output/ folder. Only xlsx to be allowed
## output filename = input_filename[_octant_analysis_mod_5000].xlsx , ie, append _octant_analysis_mod_5000 to the original filename. 

###Code

from platform import python_version
ver = python_version()

if ver == "3.8.10":
	print("Correct Version Installed")
else:
	print("Please install 3.8.10. Instruction are present in the GitHub Repo/Webmail. Url: https://pastebin.com/nvibxmjw")


mod=5000
octant_analysis(mod)






#This shall be the last lines of the code.
end_time = datetime.now()
print('Duration of Program Execution: {}'.format(end_time - start_time))
