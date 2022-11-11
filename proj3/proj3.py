import streamlit as st
st.set_page_config(layout='wide')
import os
from datetime import datetime
start_time = datetime.now()

#Help
def proj_psat_gui():
	pass


###Code

from platform import python_version
ver = python_version()

if ver == "3.8.10":
	print("Correct Version Installed")
else:
	print("Please install 3.8.10. Instruction are present in the GitHub Repo/Webmail. Url: https://pastebin.com/nvibxmjw")


proj_psat_gui()

st.title("Project 3")
st.text("Web Based Interface  using Streamlit Library by")
st.text('1. Abhay Panwar 2001CB03')
st.text('2. Jatin Kumawat 2001CB26')
st.text('It has two options~ Single File or Multi File Processing')
mod = st.number_input("Enter mod value:", format="%i", key="mod")
Correlation = st.number_input("Enter Correlation value:", format="%i", key="Correlation")
SNR = st.number_input("Enter SNR value:", format="%i", key="SNR")
Acceleration = st.number_input("Enter Acceleration value:", format="%i", key="Acceleration")
Shear_Velocity = st.number_input("Enter Shear Velocity value:", format="%i", key="Shear_Velocity")
t1, t2 = st.tabs(["Single File", "Multiple system"])
with t1:
	st.write("Single File Uploading")
	st.text('1. Upload your file.')
	st.text('2. Click on Compute Button.')
	st.write('*Not following the above steps for interacting will led to errors.')
	inputFile = st.file_uploader("Upload Excel Octant Input", type="xlsx", key="inputFile")
	reg = st.button('COMPUTE',on_click = None)
	if reg:
		
		t = datetime.now().strftime("%Y-%m-%d-%H-%M-%S")
		filename = st.session_state.inputFile.name.split(".xlsx")[0]+"_{0}_{1}.xlsx".format(st.session_state.mod2, t)
		with open("./Temp.xlsx", 'rb') as my_file:
			st.download_button(label = 'Download Output', data = my_file, file_name = filename, mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
		os.remove("./Temp.xlsx")   
	# inputFile = octant_analysis_single_file(mod, inputFile)
	# df = pandas.read_excel(inputFile)
	# st.dataframe(df)
	for i in st.session_state.keys():
		del i
import tkinter as tk
from tkinter import filedialog


			
with t2:
	st.header("MultiFile")
	st.text('1. Select the desired path.')
	st.text('2. Click on Compute Button.')
	st.write('*Not following the above steps for interacting will led to errors.')
	# Set up tkinter
	root = tk.Tk()
	root.withdraw()
	dirname = 'a'
	# Make folder picker dialog appear on top of other windows
	root.wm_attributes('-topmost', 1)
	# Folder picker button
	st.write('Please select a folder:')
	clicked = st.button('Select Folder')
	if clicked:
		dirname = st.text_input('Selected folder:', filedialog.askdirectory(master=root), key="dirname")
		
		# with st.form("my_form", clear_on_submit=False):
		# 	st.write("Inside the form")
		# 	slider_val = st.number_input("Form slider", format="%d")
		# 	checkbox_val = st.checkbox("Form checkbox")

		# 	# Every form must have a submit button.
		# 	submitted = st.form_submit_button("Submit")
		# 	if submitted:
	reg = st.button('Compute',on_click = None)
	if reg:
					file_list = glob.glob(f'{st.session_state.dirname}/*.xlsx')
					print(file_list)
					for myFile in file_list:
						octant_analysis_single_file(myFile, st.session_state.dirname,int(st.session_state.mod))
					st.write("Your Files are processed.")
					
	for i in st.session_state.keys():
		del i




#This shall be the last lines of the code.
end_time = datetime.now()
print('Duration of Program Execution: {}'.format(end_time - start_time))
