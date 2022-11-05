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

		try:
			#6 Generating the values that will be bound of intervals like 0,5000,10000 for mod = 5000 and storing them in list
			Bounds_mod_range=[]
			for count in range(0, row__, mod): #start with zero till number of rows and increment by mod
				Bounds_mod_range.append(count)
			Bounds_mod_range.append(row__+1)

			#Again making a dictionary, that will have for each octant all the intervals in form of list
			#Example {'1':[[0,4999,0], [5000,9999,0], .....], '-1':[[0,4999,0], [5000,9999,0], .....], ..........}
			#		'octant_value' : [[start_counter, end_counter, count of value]]
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

		#8 Adding mod range labels in output file
		for count in range(len(Bounds_mod_range)-1):
			#counter+2 as they will start to fill from 2nd row in Octant ID column
			#will insert the range as an string
			if(Count_range_wise['1'][count][1] != row__):
				Pointer2['Octant ID'][count+2]=str(Count_range_wise['1'][count][0])+'-'+str(Count_range_wise['1'][count][1])
				#Here border is inserted in next cell below because we are going to insert a row at 1st position which will shift all cells below by one
				_border = Myadd(_border, count+3,'Octant ID' )
			else :
				Pointer2['Octant ID'][count+2]=str(Count_range_wise['1'][count][0])+'-Last Index'
				_border = Myadd(_border, count+3,'Octant ID' )

		#9 Filling count of each octant in each mod range into csv file
		for key, value in Count_range_wise.items(): #iterating through dictionary
			for counter in range(len(Bounds_mod_range)-1): #number of ranges of mod will be constant
				Pointer2[str(key)][counter+2]=int(Count_range_wise[key][counter][2])
				#<file_handler>['<column_for_octant>'][row number] = string of Count_range_wise[key is octant][serial number of mod range][count of values]
				_border = Myadd(_border, counter+3,str(key) )
		Pointer2.rename(columns = {'Dummy':''}, inplace = True)



		#10 This dictionary Count_list contain number of times rank 1st occur for each octant
		Count_list ={'1':0, '-1':0, '2':0, '-2':0, '3':0, '-3':0, '4':0, '-4':0}

		#Calling the function Printer() to print Rank for overall count and passing it dictionary _hash, 
		#1st Row will have Rank for overall count so index = 1
		#Also receving the modified lists from function
		_border, _color = Printer(_hash, Pointer2, 1, Count_list, _border, _color)

		#11 Since we don't want to include count of Rank 1 of overall count, we redefine it
		Count_list ={'1':0, '-1':0, '2':0, '-2':0, '3':0, '-3':0, '4':0, '-4':0}

		#mod ranges will start from 2nd row
		#This index variable is used to print rank list for each mod range
		index = 2

		try:
			#12 For each mod range
			for i in range(len(Bounds_mod_range)-1):
				#Calling the Printer() Function
				# Octant ID		  +1	-1		 +2
				# Overall Count	 2610	4603	4855
				# Mod 5000			
				# 0-4999		  414	 688	815		<<<<===== Row number 2
				# 5000-9999		  380	 757	820
				# 10000-14999	  621	1016	599
				# 15000-19999	  366	682		948

				#Since Function needs a dictionary in octant:count form 
				#Directly accessing the cells for each octant for each mod range and passing them as dictionary through pre-defining as number of octant are 8 only
				#Since ranks will be printed in same row, passing (index + i) 
				#As for 2nd mod range, we will do all the operations in (index+1)th row
				_border, _color = Printer({'1':Pointer2['1'][index+i],
				'-1':Pointer2['-1'][index+i],
				'2':Pointer2['2'][index+i],
				'-2':Pointer2['-2'][index+i],
				'3':Pointer2['3'][index+i],
				'-3':Pointer2['-3'][index+i],
				'4':Pointer2['4'][index+i],
				'-4':Pointer2['-4'][index+i],
				}, Pointer2, index+i, Count_list, _border, _color)
		except ValueError:
			print('Value Error in Part 12')
		except :
			print('Other Error in Part 12')

		#Moving to three row below to the row having count of last mod range, Three row space has been left according to specification
		index +=  len(Bounds_mod_range)

		#13 There position will remain unaffected by other things so directly putting values in cells
		Pointer2['Rank Octant 4'][index] = 'Octant ID'
		_border = Myadd(_border,index +1,'Rank Octant 4')
		Pointer2['Rank Octant 4'][index+1] = '1'
		_border = Myadd(_border, index+2,'Rank Octant 4')
		Pointer2['Rank Octant 4'][index+2] = '-1'
		_border = Myadd(_border,index+3 ,'Rank Octant 4')
		Pointer2['Rank Octant 4'][index+3] = '2'
		_border = Myadd(_border, index+4,'Rank Octant 4')
		Pointer2['Rank Octant 4'][index+4] = '-2'
		_border = Myadd(_border,index+5 ,'Rank Octant 4')
		Pointer2['Rank Octant 4'][index+5] = '3'
		_border = Myadd(_border,index+6 ,'Rank Octant 4')
		Pointer2['Rank Octant 4'][index+6] = '-3'
		_border = Myadd(_border, index+7,'Rank Octant 4')
		Pointer2['Rank Octant 4'][index+7] = '4'
		_border = Myadd(_border, index+8,'Rank Octant 4')
		Pointer2['Rank Octant 4'][index+8] = '-4'
		_border = Myadd(_border, index+9,'Rank Octant 4')
		Pointer2['Rank Octant -4'][index] = 'Octant Name'
		_border = Myadd(_border,index +1,'Rank Octant -4')

		#Putting the mapping of octants
		Pointer2['Rank Octant -4'][index+1] = octant_name_id_mapping[ '1']
		_border = Myadd(_border,index+2 ,'Rank Octant -4')
		Pointer2['Rank Octant -4'][index+2] = octant_name_id_mapping[ '-1']
		_border = Myadd(_border,index+3 ,'Rank Octant -4')
		Pointer2['Rank Octant -4'][index+3] = octant_name_id_mapping[ '2']
		_border = Myadd(_border, index+4,'Rank Octant -4')
		Pointer2['Rank Octant -4'][index+4] = octant_name_id_mapping[ '-2']
		_border = Myadd(_border,index+5 ,'Rank Octant -4')
		Pointer2['Rank Octant -4'][index+5] = octant_name_id_mapping[ '3']
		_border = Myadd(_border,index+6 ,'Rank Octant -4')
		Pointer2['Rank Octant -4'][index+6] = octant_name_id_mapping[ '-3']
		_border = Myadd(_border, index+7,'Rank Octant -4')
		Pointer2['Rank Octant -4'][index+7] = octant_name_id_mapping[ '4']
		_border = Myadd(_border,index+8 ,'Rank Octant -4')
		Pointer2['Rank Octant -4'][index+8] = octant_name_id_mapping[ '-4']
		_border = Myadd(_border,index+9,'Rank Octant -4')

		#Putting the number of times Rank 1st has occured for each octant
		Pointer2['Rank1 Octant ID'][index] = 'Count of Rank 1 Mod Values'
		_border = Myadd(_border,index+1,'Rank1 Octant ID')
		Pointer2['Rank1 Octant ID'][index+1] = Count_list[ '1']
		_border = Myadd(_border,index+2 ,'Rank1 Octant ID')
		Pointer2['Rank1 Octant ID'][index+2] = Count_list[ '-1']
		_border = Myadd(_border,index+3 ,'Rank1 Octant ID')
		Pointer2['Rank1 Octant ID'][index+3] = Count_list[ '2']
		_border = Myadd(_border, index+4,'Rank1 Octant ID')
		Pointer2['Rank1 Octant ID'][index+4] = Count_list[ '-2']
		_border = Myadd(_border,index+5 ,'Rank1 Octant ID')
		Pointer2['Rank1 Octant ID'][index+5] = Count_list[ '3']
		_border = Myadd(_border, index+6,'Rank1 Octant ID')
		Pointer2['Rank1 Octant ID'][index+6] = Count_list[ '-3']
		_border = Myadd(_border,index+7 ,'Rank1 Octant ID')
		Pointer2['Rank1 Octant ID'][index+7] = Count_list[ '4']
		_border = Myadd(_border, index+8,'Rank1 Octant ID')
		Pointer2['Rank1 Octant ID'][index+8] = Count_list[ '-4']
		_border = Myadd(_border, index+9,'Rank1 Octant ID')

		#14 Storing each transition as combination using fstring like 11 for +1 to +1 and 2-1 for +2 to -1 in a dictionary
		Transition_comb =dict()
		try :
			for i in ['1','-1','2','-2','3','-3','4','-4']:
				for j in ['1','-1','2','-2','3','-3','4','-4']:
					Transition_comb[f'{i}{j}']=0 #here key is fstring and value is count of that key
		except ValueError():
			print("ValueError in Part 10")
		except :
			print("Other error in part 10")

		#15 making another dictionary such that for each transition range
		# starting bound is the key and value is the above dictionary which is Transition_comb for each range
		Transition_range_comb = dict()
		for i in range(len(Bounds_mod_range)-1):
			val = Bounds_mod_range[i]
			Transition_range_comb[val] = Transition_comb.copy()

		#16 For each counter, we will count transition
		try: 
			for counter, rows in Pointer2.iterrows():
				if counter==(row__-1): 
					continue #skip last counter as there is no row below it to make a transition
				else :
					val = mod * int(counter/mod) 
					#it is formula to find lower bound of a range to which a counter belongs
					# for 11555, counter/mod = 11555/5000 = 2.311
					# int(2.311) = 2
					# mod *2 = 5000*2 = 10,000 
					i = Pointer2['Octant'][counter]
					j = Pointer2['Octant'][counter+1]
					#using fstring, make suitable key
					Transition_range_comb[val][f'{i}{j}'] += 1
					#increasing Transition_comb[key] to keep transition count for overall
					Transition_comb[f'{i}{j}'] +=1
		except ValueError():
			print("ValueError in Part 12")
		except:
			print("Error in Part 12")
		
		#17Counting longest Subsequence for each octant
		Dict_longes_Sub_seq = { '1':[0,0,[]],'-1':[0,0,[]],'2':[0,0,[]],'-2':[0,0,[]],'3':[0,0,[]],'-3':[0,0,[]],'4':[0,0,[]],'-4':[0,0,[]]}
		# here key = Octant Value and Value  = a list which 0th index is length of longest subsequence for this octant and 
		# number of times this longest sequence has occured at 1st index and list of time interval in form of another list [start time, end time]

		index = 0
		# number of times this longest sequence has occured at 1st index
		#It should always be two less than number of rows, one due to header and other as indexing starts from zero
		#starting from first index
		index = 0
		#It should always be two less than number of rows, one due to header and other as indexing starts from zero
		while index<(row__-1):
				#Storing the octant for which, we will find the current longest subsequence
			cur = str(Pointer2['Octant'][index])
				#starts with length zero initially
			length = 0
			start_time = Pointer2['T'][index] #storing the starting time of beginning of sequence
				#keep running this while loop until we find character same as our cur octant
			while str(Pointer2['Octant'][index]) == cur and index<(row__-1):
				length += 1 #Increase current subsequence length
				index += 1 #Move to next index
				
				#If our current subsequence length is greater than previous Greatest Subsequence for current octant
			if length > Dict_longes_Sub_seq[cur][0]:
					Dict_longes_Sub_seq[cur][0] = length #Make current subsequence length as Greatest Subsequence for current octant
					Dict_longes_Sub_seq[cur][1] = 1 #It has occured first time, so count is 1
					Dict_longes_Sub_seq[cur][2].clear() #Empty the list of list as it is new beginning
					Dict_longes_Sub_seq[cur][2].append([start_time,Pointer2['T'][index-1]]) #Appending the list interval
			
				#Else If our current subsequence length is equal to previous Greatest Subsequence for current octant
			elif length == Dict_longes_Sub_seq[cur][0]:
					#Just Increment the count for this length of Subsequence
					Dict_longes_Sub_seq[cur][1] += 1
					#Appending the current time interval in the list of lists of time
					Dict_longes_Sub_seq[cur][2].append([start_time,Pointer2['T'][index-1]])


		Pointer2 = pandas.concat([pandas.DataFrame([['T','U','V','W','U Avg','V Avg','W Avg','U\'=U - U avg','V\'=V - V avg','W\'=W - W avg','Octant',None,None,None,None,None,None,None,None,None,None,None,None,None,None,None,None,None,None,None,None,None]],columns=Pointer2.columns),Pointer2],ignore_index=True)
		#18 Since we have created a new dataframe, so the index of excel needs to be reseted, it is done by below code
		Pointer2 = Pointer2.sort_index().reset_index(drop=True)
			#Now in the header row, all columns name except type 'Rank 1', 'Rank 2, and so on will become null and these will be replaced by '1', '-1', '2' and so on ....
			# Columns names changed using for loop for following names and 
			#changing others according to specification 

		Pointer2.insert(32, 'c1', '')
		Pointer2.insert(33, 'c2', '')
		Pointer2.insert(34, 'c3', '')
		Pointer2.insert(35, '1c', '')
		Pointer2.insert(36, '-1c', '')
		Pointer2.insert(37, '2c', '')
		Pointer2.insert(38, '-2c', '')
		Pointer2.insert(39, '3c', '')
		Pointer2.insert(40, '-3c', '')
		Pointer2.insert(41, '4c', '')
		Pointer2.insert(42, '-4c', '')
		counter = -1
		#Writing overall transition count table using Value_put function
		Pointer2['1c'][counter+1]='To'
		Pointer2['c2'][counter+3]='From'
		Pointer2['c3'][counter+2]='Octant #'
		_border = Myadd(_border,counter+2,'c3' )
		Pointer2['c3'][counter+3]='+1'
		_border = Myadd(_border,counter+3,'c3' )
		Pointer2['c3'][counter+4]='-1'
		_border = Myadd(_border,counter+4,'c3' )
		Pointer2['c3'][counter+5]='+2'
		_border = Myadd(_border,counter+5,'c3' )
		Pointer2['c3'][counter+6]='-2'
		_border = Myadd(_border,counter+6,'c3' )
		Pointer2['c3'][counter+7]='+3'
		_border = Myadd(_border,counter+7,'c3' )
		Pointer2['c3'][counter+8]='-3'
		_border = Myadd(_border,counter+8,'c3' )
		Pointer2['c3'][counter+9]='+4'
		_border = Myadd(_border,counter+9,'c3' )
		Pointer2['c3'][counter+10]='-4'
		_border = Myadd(_border,counter+10,'c3' )
		Pointer2['1c'][counter+2]='+1'
		_border = Myadd(_border,counter+2,'1c' )
		Pointer2['-1c'][counter+2]='-1'
		_border = Myadd(_border,counter+2,'-1c' )
		Pointer2['2c'][counter+2]='+2'
		_border = Myadd(_border,counter+2,'2c' )
		Pointer2['-2c'][counter+2]='-2'
		_border = Myadd(_border,counter+2,'-2c' )
		Pointer2['3c'][counter+2]='+3'
		_border = Myadd(_border,counter+2,'3c' )
		Pointer2['-3c'][counter+2]='-3'
		_border = Myadd(_border,counter+2,'-3c' )
		Pointer2['4c'][counter+2]='+4'
		_border = Myadd(_border,counter+2,'4c' )
		Pointer2['-4c'][counter+2]='-4'
		_border = Myadd(_border,counter+2,'-4c' )

		#Space_according_to_Octant is a dictionary which contain position for particular transition value column as they need to come in a particular order
		Space_according_to_Octant ={'1':3, '-1':4,'2':5,'-2':6,'3':7,'-3':8,'4':9, '-4':10} #{octant: position along horizontal direction}
		#19 Feeding main values in the matrix
		for i in ['1','-1','2','-2','3','-3','4','-4']:
			_max = list()
			for j in ['1','-1','2','-2','3','-3','4','-4']:
				_max.append(Transition_comb[f'{j}{i}'])
			_Great = max(_max)
			for j in ['1','-1','2','-2','3','-3','4','-4']:
				Pointer2[i+'c'][counter+Space_according_to_Octant[j]]=Transition_comb[f'{j}{i}']
				_border = Myadd(_border,counter+Space_according_to_Octant[j] , i+'c')
				if(Transition_comb[f'{j}{i}'] == _Great):
					_color = Myadd(_color,counter+Space_according_to_Octant[j] , i+'c')

		#Leaving space of rows, 
		#counter = 1 row for 'To' + 1 row for 'From' + 8 rows for 8 octants + 2 row for blank space + 1 for beginning of next matrix
		counter+=14

		try: 
			#20 Using For loop, putting each matrix in excel file
			for i in range(len(Bounds_mod_range)-1): #For each range
				val = Bounds_mod_range[i]
				#Calling Function Value_put()
				if((Bounds_mod_range[i+1]-1)!=row__):
					_border, _color = Value_put(Pointer2, f'{Bounds_mod_range[i]}-{Bounds_mod_range[i+1]-1}', counter, Transition_range_comb[val], 'Mod Transition Count', _border, _color)
				else:
					_border, _color = Value_put(Pointer2, f'{Bounds_mod_range[i]}-Last Index', counter, Transition_range_comb[val], 'Mod Transition Count', _border, _color)
				counter+=13 #Jumping to next desired location
		except TypeError():
			print("TypeError in Part 14")
		except ValueError():
			print("ValueError in Part 14")
		except:
			print("Error in Part 14")

		
		Pointer2.insert(43, 'd1', '')
		Pointer2.insert(44, 'd2', '')
		Pointer2.insert(45, 'd3', '')
		Pointer2.insert(46, 'd4', '')


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
