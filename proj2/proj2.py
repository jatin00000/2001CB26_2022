#Importing streamlit library and other necessary module
import streamlit as st
import os
import tkinter as tk
from tkinter import filedialog
#Using set_page_config(), we can control layout of page
#setting width to cover entire screen
st.set_page_config(layout='wide')

#St.markdown is used to write html in streamlit
#Hidding the warning for number_input using html display = none
st.markdown('''
<style>
div[data-testid="stMarkdownContainer"]{
	display: none;
}
</style>


''', unsafe_allow_html = True)
from datetime import datetime
start_time = datetime.now()

#Help

#Preventing display of warning in terminal using Warning module
import warnings
warnings.simplefilter(action='ignore', category= (FutureWarning, UserWarning))
from platform import python_version
ver = python_version()

if ver == "3.8.10":
	print("Correct Version Installed")
else:
	print("Please install 3.8.10. Instruction are present in the GitHub Repo/Webmail. Url: https://pastebin.com/nvibxmjw")


def proj_octant_gui():
	#Printing title on screen
	st.title("Project 2")
	st.header("GUI version of tut07")

	#Printing little bit info
	#st.text(<text>) is used to print text on screen
	st.text("Web Based Interface for Tut07 using Streamlit Library by")
	st.text('1. Abhay Panwar 2001CB03')
	st.text('2. Jatin Kumawat 2001CB26')
	st.text('It has two options~ Single File or Multi File Processing')

	#st.tabs(<list of name of string>) is used to create tabs
	t1, t2 = st.tabs(["Single File", "Multiple system"])

	#Defining tab t1
	with t1:
		#Text to be displayed
		st.write("Single File Uploading")
		st.text('1. Upload your file.')
		st.text('2. Enter mod value.')
		st.text('3 Click on Compute Button.')
		st.write('*Not following the above steps for interacting will led to errors.')

		#st.file_uploader(<text to be displayed>, type of file, key for this element) is use to take specific type file input from user
		inputFile = st.file_uploader("Upload Excel Octant Input", type="xlsx", key="inputFile")

		#taking input for mod value
		#st.number_input is use to take number input, 'format' is use to specify type of number like decimal, integer
		mod = st.number_input("Enter mod value:", format="%i", key="mod2")

		#Creating a compute button
		reg = st.button('COMPUTE',on_click = None)
		#If button is clicked
		if reg:
			#In streamlit, on clicking a button
			#The script is refreshed and all the run time values like here of mod, file input are lost
			#using st.session_state is a dictionary type use to store values in run time
			#When the script is refreshed due to button click

			#This session state dictionary data is not lost and we can use these values in program


			#Calling this function designed to process a single file
			octant_analysis_For_single_file(st.session_state.inputFile, int(st.session_state.mod2))

			#Taking current time in format 'Year-month-day-hour-minute-second'
			t = datetime.now().strftime("%Y-%m-%d-%H-%M-%S")

			#generating name for output file using fstring
			filename = st.session_state.inputFile.name.split(".xlsx")[0]+"_{0}_{1}.xlsx".format(st.session_state.mod2, t)

			#opening Temp.xlsx file generated during octant_analysis_For_single_file()
			with open("./Temp.xlsx", 'rb') as my_file:
				#Providing option to download using download button of streamlit
				#label is text for button
				#data is the dataframe
				#file_name is the name of downloading file
				#mime is the type of file
				st.download_button(label = 'Download Output', data = my_file, file_name = filename, mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
			
			#Deleting the file from current directory
			os.remove("./Temp.xlsx")  

		#Clearing the session state as it could get clear now	 
		for i in st.session_state.keys():
			del i



	#Defining second Tab		
	with t2:
		#Text to be displayed
		st.header("MultiFile")
		st.text('1. Select the desired path.')
		st.text('2. Enter mod value.')
		st.text('3 Click on Compute Button.')
		st.write('*Not following the above steps for interacting will led to errors.')

		#using tkinter module here to take input a path of folder
		# Set up tkinter

		#Tk() helps to display the root window and manages all the other components of the tkinter application
		root = tk.Tk()

		# Tkinter withdraw method hides the window without destroying it internally.
		root.withdraw()

		dirname = 'a'
		# Make folder picker dialog appear on top of other windows
		root.wm_attributes('-topmost', 1)

		# Folder picker button
		st.write('Please select a folder:')
		clicked = st.button('Select Folder')
		if clicked:
			#st.text_input is a input taker in simple text format of streamlit
			#filedialog.askdirector() is provides a unique way to select file
			#And is also select path
			dirname = st.text_input('Selected folder:', filedialog.askdirectory(master=root), key="dirname")

		#taking input for mod value
		#st.number_input is use to take number input, 'format' is use to specify type of number like decimal, integer
		mod = st.number_input("Enter mod value:", format="%d", key="mod")	

		#Creating a compute button
		reg = st.button('Compute',on_click = None)
		#If button is clicked
		if reg:
						
						#used to return all file paths that match a specific pattern. 
						file_list = glob.glob(f'{st.session_state.dirname}/*.xlsx')
						for myFile in file_list:
							#Taking each file, one by one
							octant_analysis_single_file(myFile, st.session_state.dirname,int(st.session_state.mod))
						st.write("Your Files are processed.")

		#Clearing the session state as it could get clear now			
		for i in st.session_state.keys():
			del i

proj_octant_gui()




#This shall be the last lines of the code.
end_time = datetime.now()
print('Duration of Program Execution: {}'.format(end_time - start_time))
