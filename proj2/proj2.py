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


		
		
		
		#24 Opening in openpyxl module to do formating
		# Method to save in specific folder: os.path.join method, for joining one or more path components.
		# a = 1.0.xlsx
		# to get only "1.0" split by ".xlsx" and take first part which is the name

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

	

proj_octant_gui()




#This shall be the last lines of the code.
end_time = datetime.now()
print('Duration of Program Execution: {}'.format(end_time - start_time))
