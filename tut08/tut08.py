import re
from datetime import datetime
start_time = datetime.now()

#Help
def scorecard():
	#A dictionary having name of each player and short name used in commmentary
	name = {'Bhuvneshwar':'Bhuvneshwar Kumar','Rahul':'KL Rahul', 'Rizwan':'Mohammad Rizwan(w)', 'Babar Azam':'Babar Azam(c)', 'Arshdeep Singh':'Arshdeep Singh', 'Fakhar Zaman':'Fakhar Zaman', 'Hardik Pandya':'Hardik Pandya', 'Avesh Khan':'Avesh Khan', 'Iftikhar Ahmed':'Iftikhar Ahmed', 'Chahal':'Yuzvendra Chahal', 'Jadeja':'Ravindra Jadeja', 'Khushdil':'Khushdil Shah', 'Shadab Khan':'Shadab Khan', 'Asif Ali':'Asif Ali', 'Mohammad Nawaz':'Mohammad Nawaz', 'Haris Rauf':'Haris Rauf', 'Naseem Shah':'Naseem Shah', 'Dahani':'Shahnawaz Dahani', 'Rohit':'Rohit Sharma(c)', 'Kohli':'Virat Kohli', 'Suryakumar Yadav':' Suryakumar Yadav', 'Karthik':'Dinesh Karthik(w)'}
	#List of names of players of Team India
	TeamIndia =[ 'Rohit','Rahul' ,'Kohli','Suryakumar Yadav', 'Karthik', 'Hardik Pandya', 'Jadeja', 'Bhuvneshwar', 'Avesh Khan', 'Chahal', 'Arshdeep Singh']
	TeamPak = ['Babar Azam', 'Rizwan', 'Fakhar Zaman', 'Iftikhar Ahmed', 'Khushdil', 'Asif Ali', 'Shadab Khan', 'Mohammad Nawaz', 'Naseem Shah', 'Haris Rauf', 'Dahani']
	#Opening the file
	f = open('Scorecard.txt', 'w')
	#Reading the innings file
	f1 = open('pak_inns1.txt', 'r')
	f2 = open('india_inns2.txt', 'r')

	#Reading all the lines
	Pak = f1.readlines()
	Ind = f2.readlines()

	#Storing the number of balls thrown and run scored by each player in a dictionary 
	#key as player short name and value as list
	#Initializing
	score = dict()
	for i in name:
		score[i] = [0,0]
	
	#Reaing Pakistan Innings
	for i in Pak:
		#If it contains sentences other than new line character
		if(len(i.split(","))>1):
			#taking the first part having
			s1 = i.split(",")[0]
			#Searching for name of bowler
			#Pattern is it should be in between number 14.2 and 'to' 
			#The below expression returns all such string
			searchObj = re.search( r'(?:\d*\.\d+|\d+)(?: & b)? (.+?)(?: to |$)', s1, re.M|re.I)
			bowler = searchObj.group(1)

			#The below expression returns all string from 'to' till end which is name of batsman
			searchObj = re.search( r'to(?: & b)? (.+)', s1, re.M|re.I)
			batsman = searchObj.group(1)
			run = 0

			#Next string is for number of runs
			s2 = i.split(",")[1].strip()
			#doing stripping through space

			#if run 1,2,3 is scored, then converting to int and increasing the number of run
			if(len(s2.split(" "))>1 and  s2.split(" ")[0].isnumeric()):
				run = int(s2.split(" ")[0])
			#And if it is a single expression, and 'FOUR' means 4 runs and 'SIX' means 6 runs scored
			elif(s2.split(" ")[0] == 'FOUR'):
				run = 4
			elif(s2.split(" ")[0] == 'SIX'):
				run = 6

			#Increasing amount of balls for bowler and runs for batsman
			score[bowler][0] += 1
			score[batsman][1] += run	

	
###Code

from platform import python_version
ver = python_version()

if ver == "3.8.10":
	print("Correct Version Installed")
else:
	print("Please install 3.8.10. Instruction are present in the GitHub Repo/Webmail. Url: https://pastebin.com/nvibxmjw")


scorecard()






#This shall be the last lines of the code.
end_time = datetime.now()
print('Duration of Program Execution: {}'.format(end_time - start_time))
