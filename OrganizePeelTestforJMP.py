import glob
import os
import string
import os.path
import sys
import shutil
import xlrd
import csv
import argparse
import xlsxwriter
import pptx
#import antigravity
import pandas as pd
import numpy as np
import matplotlib as plt
import seaborn as sns
import tkinter as tk
from tkinter.filedialog import *
from tkinter.messagebox import *
from pptx import Presentation
from pptx.util import Inches
from datetime import date
from pathlib import Path


# Part 1: Define useful script specific functions here.


def organizepeeltest():

	savename = savenameentry.get()

	location = askdirectory(title = "Select Location of Pull Test Data.")
	if(not location):
		sys.exit(0)

	cleandatalocation = askdirectory(title = "Select Location to Save Cleaned Data.")
	if(not cleandatalocation):
		sys.exit(0)

	os.chdir(location)

	filenames = glob.glob("*.csv")

	col_list = ["Peel displacement", "Force", "Force / Width"]

	df_list = []

	for file in filenames:

		os.chdir(location)
		
		df = pd.read_csv(file, header = 4, usecols = col_list)

		#Drop/rename some rows & columns
		df = df.drop(0)

		df = df.rename(columns = {"Peel displacement":"Peel Displacement (mm)", "Force":"Force (N)", "Force / Width":"Force/Width (N/mm)"})

		file_noext = (file.rsplit(".", 1))[0]

		#Create new column to indicate sample number
		samplenamelist = []

		for x in range(len(df)):
			samplenamelist.append(str(file_noext))

		df.insert(0, "Sample", samplenamelist, True)

		os.chdir(cleandatalocation)

		#Export cleaned CSV files individually
		df.to_csv(str(file_noext)+'_Clean.csv')

		df_list.append(df)

	os.chdir(cleandatalocation)

	extension = 'csv'
	all_filenames = [i for i in glob.glob('*.{}'.format(extension))]

	#Combine all files in the list
	combined_csv = pd.concat([pd.read_csv(f) for f in all_filenames])

	#Export to csv
	combined_csv.to_csv(str(savename)+".csv", index = False, encoding = 'utf-8-sig')


# Part 2: Create the GUI.


peeltestgui = tk.Tk()


# Create homepage screen.


peeltestcanvas = tk.Canvas(peeltestgui, width = 400, height = 220)
peeltestcanvas.pack()


# Create boxes to take user input.


savenameentry = tk.Entry(peeltestgui)
peeltestcanvas.create_window(200, 75, window = savenameentry)

savenamelabel = tk.Label(text = "Peel Test Combinred CSV Name")
peeltestcanvas.create_window(200, 50, window = savenamelabel)


# Create buttons to select powerpoint formatting.


organizepeeltestbutton = tk.Button(text = "Organize Peel Test Data", command = organizepeeltest)
peeltestcanvas.create_window(200, 130, window = organizepeeltestbutton)


peeltestgui.mainloop()