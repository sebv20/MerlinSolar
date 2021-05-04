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


# Part 1: How to import all the files we want


# Define random useful functions here


# Precondition: s is a single character string.
# Postcondition: float(s) is the single character string changed to a float and s is the single character string returned as a string.


def maybe_float(s):
	try:
		return float(s)
	except (ValueError, TypeError):
		return s


# Precondition: extensionlist is a full list of filenames containing the extension (.jpg, .png, .heic, etc...) at the end.
# Postcondition: filenames is a full list of filenames without the extension at the end.


def stem_list(extensionlist):
	noextensionlist = []

	for file in extensionlist:
		y = Path(file).stem
		noextensionlist.append(y)

	return noextensionlist


# Define useful script specific functions here


# Precondition: function is tied to a button input.
# Postcondition: function will take all the images from all the subdirectories within a directory and add them all onto a different slide per subdirectory.


def sortcondition():

	location = askdirectory(title = "Select post condition folder.")
	if(not location):
		sys.exit(0)

	splitlocationpath = os.path.split(location)
	postcondition = splitlocationpath[1]

	condition = os.scandir(location)

	savename = savenameentry.get()

	pptx_location = askopenfilename(filetypes = [("Powerpoint Files", "*.pptx")], title = "Select Powerpoint Template File.")
	if not(pptx_location):
		sys.exit(0)

	prs = Presentation(pptx_location)

	for folder in condition:
		if folder.is_dir():
			imagefiles = []

			os.chdir(folder)
			filenames = glob.glob("*.jpg")

			for file in filenames:
				imagefiles.append(file)

			addimages(folder, postcondition, pptx_location, prs, imagefiles)

	save_location = askdirectory(title = "Select location to save presentation.")
	os.chdir(save_location)
	prs.save(str(savename)+".pptx")


# Precondition: pptx_location is the location of the powerpoint file template. prs is the powerpoint presentation. modulelist is the list of images from the given subdirectory.
# Postcondition: function will all all the images from a certain directory to a new slide.


def addimages(folder, postcondition, pptx_location, prs, modulelist):

	lyt_1 = prs.slide_layouts[0]
	lyt_2 = prs.slide_layouts[1]
	lyt_3 = prs.slide_layouts[2]
	lyt_4 = prs.slide_layouts[3]
	lyt_5 = prs.slide_layouts[4]
	lyt_6 = prs.slide_layouts[5]
	lyt_7 = prs.slide_layouts[6]
	lyt_8 = prs.slide_layouts[7]

	pictureslides = prs.slides.add_slide(lyt_5)
	pictureslidestitle = pictureslides.shapes.title

	newfolder = str(folder)
	foldername = newfolder[newfolder.find("'")+1:newfolder.rfind("'")]

	pictureslidestitle.text = str(foldername)+" @ "+str(postcondition)

	width_full_standard_slide = Inches(10)
	width_full_widescreen_slide = Inches(13.333)
	height_full_slide = Inches(7.5)
	top_margin = Inches(1.5)
	side_margin = Inches(0.1)
	margin_in_between = Inches(0.5)

	top_6_1 = top_margin
	left_6_1 = side_margin
	top_6_2 = (height_full_slide/2)+(1.5*margin_in_between)
	left_6_2 = (width_full_widescreen_slide/3)+left_6_1
	left_6_3 = (2*(width_full_widescreen_slide/3))+left_6_1
	width_6 = (width_full_widescreen_slide/3)-(2*left_6_1)
	height_6 = (height_full_slide/2)-top_6_1

	for image in modulelist:
		pictureslides.shapes.add_picture(image, left_6_2, top_6_1)


# Part 3: Create the GUI


powerpointgui = tk.Tk()


# Create homepage screen.


powerpointcanvas = tk.Canvas(powerpointgui, width = 400, height = 220)
powerpointcanvas.pack()


# Create boxes to take user input.


savenameentry = tk.Entry(powerpointgui)
powerpointcanvas.create_window(200, 90, window = savenameentry)

savenamelabel = tk.Label(text = "Powerpoint File Name")
powerpointcanvas.create_window(200, 50, window = savenamelabel)


# Create buttons to select powerpoint formatting.


sortconditionbutton = tk.Button(text = "Format Powerpoint by Post Condition", command = sortcondition)
powerpointcanvas.create_window(200, 130, window = sortconditionbutton)

powerpointgui.mainloop()