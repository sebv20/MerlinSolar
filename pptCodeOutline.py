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


# Define random useful functions here


# Function Purpose: function changes a single character into a float.
# Precondition: s is a single character string.
# Postcondition: float(s) is the single character string changed to a float and s is the single character string returned as a string.


def maybe_float(s):
	try:
		return float(s)
	except (ValueError, TypeError):
		return s


# Function Purpose: function removes the extension at the end of a filename.
# Precondition: extensionlist is a full list of filenames containing the extension (.jpg, .png, .heic, etc...) at the end.
# Postcondition: filenames is a full list of filenames without the extension at the end.


def stem_list(extensionlist):
	noextensionlist = []

	for file in extensionlist:
		y = Path(file).stem
		noextensionlist.append(y)

	return noextensionlist


# Part 1: How to import all the files we want



# Define useful script specific functions here

# Precondition: filenames is the full list of image filenames as strings. location is an integer representing the site where the panel was manufactured.
# Postcondition: a list is returned containing only the filenames for panels manufactured at that site.

# def sortlocation(filenames, location):
# 	imagefiles = []
# 	errorfiles = []
# 	for file in filenames:
# 		original_characters = list(file)
# 		characters = [maybe_float(v) for v in original_characters]
# 		if type(characters[0]) == str and type(characters[1]) == float and location == 1:
# 			imagefiles.append(file)
# 		elif type(characters[0]) == float and type(characters[1]) == float and location == 2:
# 			imagefiles.append(file)
# 		elif type(characters[0]) == str and type(characters[1]) == str and location == 3:
# 			imagefiles.append(file)
# 		else:
# 			errorfiles.append(file)
# 	print(errorfiles)

# 	return imagefiles


# Precondition: Filenames is a full list of image filenames as strings. Criteria is a user inputted string representing the beginning condition we want the filename to match.
# Postcondition: A list is returned containing only the filenames for modules matching the given criteria.


def sortcriteria(filenames, criteria):
	imagefiles = []
	errorfiles = []

	for file in filenames:
		if len(file[:file.rfind("-")]) <= 16:
			firstpart = file[:file.rfind("-")]
			if firstpart == criteria:
				imagefiles.append(file)
			else:
				errorfiles.append(file)
		elif len(file[:file.rfind("-")]) >= 16:
			firstpart = file[:(file[:file.rfind("-")]).rfind("-")]
			if firstpart == criteria:
				imagefiles.append(file)
			else:
				errorfiles.append(file)


	print("The following files don't fit the given criteria:", errorfiles)

	return imagefiles


def sortdate(filenames, year, month, day):
	imagefiles = []
	errorfiles = []

	yeardict = {

		"2020": "A",
		"2021": "B",
		"2022": "C",
		"2023": "D",
		"2024": "E",
		"2025": "F",
		"2026": "G",
		"2027": "H",
		"2028": "I",
		"2029": "J",
		"2030": "K",
		"2031": "L"

	}

	monthdict = {

		"january": "A",
		"february": "B",
		"march": "C",
		"april": "D",
		"may": "E",
		"june": "F",
		"july": "G",
		"august": "H",
		"september": "I",
		"october": "J",
		"november": "K",
		"december": "L"

	}

	daydict = {

		"1": "1",
		"2": "2",
		"3": "3",
		"4": "4",
		"5": "5",
		"6": "6",
		"7": "7",
		"8": "8",
		"9": "9",
		"10": "A",
		"11": "B",
		"12": "C",
		"13": "D",
		"14": "E",
		"15": "F",
		"16": "G",
		"17": "H",
		"18": "J",
		"19": "K",
		"20": "L",
		"21": "M",
		"22": "N",
		"23": "P",
		"24": "R",
		"25": "S",
		"26": "T",
		"27": "U",
		"28": "V",
		"29": "W",
		"30": "X",
		"31": "Y"

	}

	for file in filenames:
		if len(file[:file.rfind("-")]) <= 16:
			itemcode = file[file.rfind("-")+1:]
			characters = list(itemcode)
			if characters[0] == yeardict.get(year.lower()) and characters[1] == monthdict.get(month.lower()) and characters[2] == daydict.get(day.lower()):
				imagefiles.append(file)
			else:
				errorfiles.append(file)
		elif len(file[:file.rfind("-")]) >= 16:
			itemcode = file[((file[:file.rfind("-")]).rfind("-"))+1:file.rfind("-")]
			characters = list(itemcode)
			if characters[0] == yeardict.get(year.lower()) and characters[1] == monthdict.get(month.lower()) and characters[2] == daydict.get(day.lower()):
				imagefiles.append(file)
			else:
				errorfiles.append(file)

	print("The following files don't fit the given date:", errorfiles)

	return imagefiles


# Precondition: Filenames is a full list of image filenames as strings. Celltype is a user inputted string representing the cell type of the module(s).
# Postcondition: A list is returned containing only the filenames for modules with the desired cell type.


def sortcelltype(filenames, celltype):
	imagefiles = []
	errorfiles = []

	celltypedict = {

	"parent": "_",
	"mono m2 eepv": "A",
	"perc m2 eepv": "B",
	"perc g1 ure": "C",
	"hjt m2 + hevel": "D",
	"hjt m2 kaneka": "E",
	"perc m2 ure": "F",
	"engineering": "Z"

	}

	for file in filenames:
		if len(file[:file.rfind("-")]) <= 16:
			itemcode = file[file.rfind("-")+1:]
			characters = list(itemcode)
			if characters[3] == celltypedict.get(celltype.lower()):
				imagefiles.append(file)
			else:
				errorfiles.append(file)
		elif len(file[:file.rfind("-")]) >= 16:
			itemcode = file[((file[:file.rfind("-")]).rfind("-"))+1:file.rfind("-")]
			characters = list(itemcode)
			if characters[3] == celltypedict.get(celltype.lower()):
				imagefiles.append(file)
			else:
				errorfiles.append(file)

	print("The following files don't fit the given cell type:", errorfiles)

	return imagefiles


# Precondition: Filenames is a full list of image filenames as strings. Encap is a user inputted string representing the encapsulant(s) used in the module(s).
# Postcondition: A list is returned containing only the filenames for modules with the desired encapsulant type.


def sortencap(filenames, encap):
	imagefiles = []
	errorfiles = []

	encapdict = {

	"parent": "_",
	"mitsui eva str eva": "A",
	"cybrid poe cybrid 0.2 poe": "B",
	"mitsui poe cybrid 0.2 poe": "C",
	"first eva str eva 0.2": "D",
	"first eva 0.45mm": "E",
	"cybrid poe 0.45 str eva 0.2": "G",
	"provisional": "F",
	"engineering": "Z"

	}

	for file in filenames:
		if len(file[:file.rfind("-")]) <= 16:
			itemcode = file[file.rfind("-")+1:]
			characters = list(itemcode)
			if characters[4] == encapdict.get(encap.lower()):
				imagefiles.append(file)
			else:
				errorfiles.append(file)
		elif len(file[:file.rfind("-")]) >= 16:
			itemcode = file[((file[:file.rfind("-")]).rfind("-"))+1:file.rfind("-")]
			characters = list(itemcode)
			if characters[4] == encapdict.get(encap.lower()):
				imagefiles.append(file)
			else:
				errorfiles.append(file)

	print("The following files don't fit the given encapsulant:", errorfiles)

	return imagefiles


# Precondition: Filenames is a full list of image filenames as strings. Gridversion is a user inputted string representing the grid versions used in the module(s).
# Postcondition: A list is returned containing only the filenames for modules with the desired grid version.


def sortgridversion(filenames, gridversion):
	imagefiles = []
	errorfiles = []

	gridversiondict = {

	"por": "1",
	"gen 2 black": "2",
	"gen 2 copper": "3",
	"engineering": "0"

	}

	for file in filenames:
		if len(file[:file.rfind("-")]) <= 16:
			itemcode = file[file.rfind("-")+1:]
			characters = list(itemcode)
			if characters[5] == gridversiondict.get(gridversion.lower()):
				imagefiles.append(file)
			else:
				errorfiles.append(file)
		elif len(file[:file.rfind("-")]) >= 16:
			itemcode = file[((file[:file.rfind("-")]).rfind("-"))+1:file.rfind("-")]
			characters = list(itemcode)
			if characters[5] == gridversiondict.get(gridversion.lower()):
				imagefiles.append(file)
			else:
				errorfiles.append(file)

	print("The following files don't fit the given grid version:", errorfiles)

	return imagefiles


# Precondition: Filenames is a full list of image filenames as strings. Mfglocation is a user inputted string representing the manufacturing location of the module(s).
# Postcondition: A list is returned containing only the filenames for modules with the desired manufacturing location.


def sortmfglocation(filenames, mfglocation):
	imagefiles = []
	errorfiles = []

	mfglocationdict = {

	"san jose": "1",
	"waaree": "2",
	"laguna": "3"

	}

	for file in filenames:
		if len(file[:file.rfind("-")]) <= 16:
			itemcode = file[file.rfind("-")+1:]
			characters = list(itemcode)
			if characters[6] == mfglocationdict.get(mfglocation.lower()):
				imagefiles.append(file)
			else:
				errorfiles.append(file)
		elif len(file[:file.rfind("-")]) >= 16:
			itemcode = file[((file[:file.rfind("-")]).rfind("-"))+1:file.rfind("-")]
			characters = list(itemcode)
			if characters[6] == mfglocationdict.get(mfglocation.lower()):
				imagefiles.append(file)
			else:
				errorfiles.append(file)

	print("The following files don't fit the given manufacturing location:", errorfiles)

	return imagefiles


# Precondition: Filenames is a full list of image filenames as strings. Family is a user inputted string representing the family of the module(s).
# Postcondition: A list is returned containing only the filenames for modules with the desired family.


def sortfamily(filenames, family):
	imagefiles = []
	errorfiles = []

	familydict = {

	"fx transportation": "F",
	"fx roofing": "R",
	"back rail glass roofing": "B",
	"fx marine": "M",
	"nishati": "N",
	"honeycomb board": "H",
	"gx": "G",
	"fx portables/folding": "X",
	"specialty": "S"

	}

	for file in filenames:
		if len(file[:file.rfind("-")]) <= 16:
			itemcode = file[file.rfind("-")+1:]
			characters = list(itemcode)
			if characters[7] == familydict.get(family.lower()):
				imagefiles.append(file)
			else:
				errorfiles.append(file)
		elif len(file[:file.rfind("-")]) >= 16:
			itemcode = file[((file[:file.rfind("-")]).rfind("-"))+1:file.rfind("-")]
			characters = list(itemcode)
			if characters[7] == familydict.get(family.lower()):
				imagefiles.append(file)
			else:
				errorfiles.append(file)

	print("The following files don't fit the given family:", errorfiles)

	return imagefiles


# Precondition: Filenames is a full list of image filenames as strings. Color is a user inputted string representing the color of the module(s).
# Postcondition: A list is returned containing only the filenames for modules with the desired color.


def sortcolor(filenames, color):
	imagefiles = []
	errorfiles = []

	colordict = {

	"black": "B",
	"white": "W",
	"transparent": "T",
	"camo": "C",
	"multicolor": "M",
	"engineering": "X"

	}

	for file in filenames:
		if len(file[:file.rfind("-")]) <= 16:
			itemcode = file[file.rfind("-")+1:]
			characters = list(itemcode)
			if characters[8] == colordict.get(color.lower()):
				imagefiles.append(file)
			else:
				errorfiles.append(file)
		elif len(file[:file.rfind("-")]) >= 16:
			itemcode = file[((file[:file.rfind("-")]).rfind("-"))+1:file.rfind("-")]
			characters = list(itemcode)
			if characters[8] == colordict.get(color.lower()):
				imagefiles.append(file)
			else:
				errorfiles.append(file)

	print("The following files don't fit the given color:", errorfiles)

	return imagefiles


# Precondition: Filenames is a full list of image filenames as strings. Cellcut is a user inputted string representing the cell size used in the module(s).
# Postcondition: A list is returned containing only the filenames for modules with the desired cell size.


def sortcellcut(filenames, cellcut):
	imagefiles = []
	errorfiles = []

	cellcutdict = {

	"full": "F",
	"half": "H",
	"quarter": "Q",
	"eighth": "E",
	"sixteenth": "S"

	}

	for file in filenames:
		if len(file[:file.rfind("-")]) <= 16:
			itemcode = file[file.rfind("-")+1:]
			characters = list(itemcode)
			if characters[9] == cellcutdict.get(cellcut.lower()):
				imagefiles.append(file)
			else:
				errorfiles.append(file)
		elif len(file[:file.rfind("-")]) >= 16:
			itemcode = file[((file[:file.rfind("-")]).rfind("-"))+1:file.rfind("-")]
			characters = list(itemcode)
			if characters[9] == cellcutdict.get(cellcut.lower()):
				imagefiles.append(file)
			else:
				errorfiles.append(file)

	print("The following files don't fit the given cell cut:", errorfiles)

	return imagefiles


# Precondition: Filenames is a full list of image filenames as strings. Arrayconfiguration is a user inputted string representing the array configuration of the module(s).
# Postcondition: A list is returned containing only the filenames for modules with the desired array configuration.


def sortarrayconfig(filenames, arrayconfiguration):
	imagefiles = []
	errorfiles = []

	sixteenthcellconfigdict = {

	"10 sc config (trucklite)": "AA"

	}

	quartercellconfigdict = {

	"3x(6x6) qc xp config": "AA",
	"4x14 qc config": "AB",
	"4x8 qc config": "AC",
	"8x4 qc config": "AD",
	"12x6 qc config": "AE",
	"14x6 qc config": "AF"

	}

	halfcellconfigdict = {

	"3x(4x5) hc multifold (nishati)": "AA",
	"8x6 hc config": "AB",
	"4x12 hc config": "AC",
	"4x11 hc config": "AD",
	"3x(2x7) hc bxd config": "AE",
	"6x7-1 hc config": "AF",
	"6x6 hc horizontal": "AG",
	"3x(4x3) hc xp config": "AH",
	"5x7 hc config": "AI",
	"2x(4x4) hc multifold (nishati)": "AJ",
	"4x(2x3) hc minixp config": "AK",
	"5x4 hc config - king": "AL",
	"4x(2x2) hc minixp config": "AM",
	"3x6 hc config": "AN",
	"4x6-1 hc config": "AO",
	"4x6 hc config": "AP",
	"6x4 hc config": "AQ",
	"6x7 hc config": "AR"

	}

	fullcellconfigdict = {

	"6x12 fc br config": "AA",
	"6x12 fc s config": "AB",
	"6x12 fc p,corner jbox, turnt": "AC",
	"6x12 fc p,corner jbox, standard": "AD",
	"8x6 fc config": "AE",
	"4x12 fc config": "AF",
	"3x14-1 fc config": "AG",
	"3x(4x3) fc xp config": "AH",
	"3x12 fc config": "AI",
	"2x18 fc config": "AJ",
	"6x6 fc config": "AK",
	"4x9 fc config": "AL",
	"4x8 fc config": "AM",
	"2x(4x4) fc multifold (nishati)": "AN",
	"2x15 fc config": "AO",
	"3x(4x2) fc trifold nishati": "AP",
	"2x12 fc config": "AQ",
	"3x8 fc config": "AR",
	"4x6 fc config": "AS",
	"3x7 fc config": "AT",
	"2x9 fc config": "AU",
	"2x8 fc config": "AV",
	"4x4 fc config": "AW",
	"4x3 fc config": "AX",
	"2x6 fc config": "AY",
	"2x3 fc config": "AZ",
	"12x8 fc tinywatts": "BA",
	"8x6 fc navistar": "BB",
	"14x6 fc config": "BC"

	}

	for file in filenames:
		if len(file[:file.rfind("-")]) <= 16:
			itemcode = file[file.rfind("-")+1:]
			characters = list(itemcode)
			if characters[10]+characters[11] == sixteenthcellconfigdict.get(arrayconfiguration.lower()):
				imagefiles.append(file)
			elif characters[10]+characters[11] == quartercellconfigdict.get(arrayconfiguration.lower()):
				imagefiles.append(file)
			elif characters[10]+characters[11] == halfcellconfigdict.get(arrayconfiguration.lower()):
				imagefiles.append(file)
			elif characters[10]+characters[11] == fullcellconfigdict.get(arrayconfiguration.lower()):
				imagefiles.append(file)
			else:
				errorfiles.append(file)
		elif len(file[:file.rfind("-")]) >= 16:
			itemcode = file[((file[:file.rfind("-")]).rfind("-"))+1:file.rfind("-")]
			characters = list(itemcode)
			if characters[10]+characters[11] == sixteenthcellconfigdict.get(arrayconfiguration.lower()):
				imagefiles.append(file)
			elif characters[10]+characters[11] == quartercellconfigdict.get(arrayconfiguration.lower()):
				imagefiles.append(file)
			elif characters[10]+characters[11] == halfcellconfigdict.get(arrayconfiguration.lower()):
				imagefiles.append(file)
			elif characters[10]+characters[11] == fullcellconfigdict.get(arrayconfiguration.lower()):
				imagefiles.append(file)
			else:
				errorfiles.append(file)

	print("The following files don't fit the given array configuration:", errorfiles)

	return imagefiles


# Precondition: Filenames is a full list of image filenames as strings. Serialnumber is a user inputted string representing the serial number of the module(s).
# Postcondition: A list is returned containing only the filenames for modules with the desired serial number.


def sortserialnum(filenames, serialnumber):
	imagefiles = []
	errorfiles = []

	for file in filenames:
		if len(file[:file.rfind("-")]) <= 16:
			itemcode = file[file.rfind("-")+1:]
			characters = list(itemcode)
			if characters[12]+characters[13]+characters[14]+characters[15] == serialnumber.lower():
				imagefiles.append(file)
			else:
				errorfiles.append(file)
		elif len(file[:file.rfind("-")]) >= 16:
			itemcode = file[((file[:file.rfind("-")]).rfind("-"))+1:file.rfind("-")]
			characters = list(itemcode)
			if characters[12]+characters[13]+characters[14]+characters[15] == serialnumber.lower():
				imagefiles.append(file)
			else:
				errorfiles.append(file)

	print("The following files don't fit the given serial number:", errorfiles)

	return imagefiles


# Precondition: Filenames is a full list of image filenames as strings. Condition is a user inputted string representing the condition of the module(s) at the time the photo(s) were taken.
# Postcondition: A list is returned containing only the filenames for modules with the desired condition.


def sortcondition(filenames, condition):
	imagefiles = []
	errorfiles = []

	for file in filenames:
		if len(file[:file.rfind("-")]) >= 16:
			postcondition = file[file.rfind("-")+1:file.rfind(".")]
			if str(postcondition) == str(condition):
				imagefiles.append(file)
			else:
				errorfiles.append(file)
		elif len(file[:file.rfind("-")]) <= 16:
			errorfiles.append(file)

	print("The following files don't fit the given postcondition:", errorfiles)

	return imagefiles


# Function purpose is to sort images into the powerpoint with each slide being of modules with different pre criterias.
# Precondition: criterialist is a full list of filenames with the same pre criteria as strings.
# Postcondition: the function adds a slide and inserts a slide and inserts all the images of the given pre criteria into the slide.


def formatbycriteria():
	#Creating this list just in case the number of images per slide will be more than 9 images.
		#POSSIBLE IMPROVEMENT AREA: Find some way to let the user know (maybe a pop up screen) when this happens.
	listtoobig = []

	#The user selected pre criteria from the GUI.
	criteria = criteriaentry.get()

	#Prompt the user to select the directory that contains all the possible image files they want to filter from.
	location = askdirectory(title = "Select location of image files")
	if(not location):
		sys.exit(0)

	#Change the current directory to the selected location.
	os.chdir(location)
	#Get a list of all files in the current directory that are jpeg images.
	filenames = glob.glob("*.jpg")

	#Find out how many photos have a module ID that matches the desired pre criteria.
	criterialist = sortcriteria(filenames, str(criteria))

	#Select the right number of images to insert per slide.
	if len(criterialist) == 0:
		print("There are no images that fit the given criteria.")
	elif len(criterialist) == 1:
		addoneimage(criterialist)
	elif len(criterialist) == 2:
		addtwoimage(criterialist)
	elif len(criterialist) == 3:
		addthreeimage(criterialist)
	elif len(criterialist) == 4:
		addfourimage(criterialist)
	elif len(criterialist) == 5:
		addfiveimage(criterialist)
	elif len(criterialist) == 6:
		addsiximage(criterialist)
	elif len(criterialist) == 7:
		addsevenimage(criterialist)
	elif len(criterialist) == 8:
		addeightimage(criterialist)
	else:
		listtoobig.append(file)

	#Currently the only indication that the user has tried putting more than 9 images per slide.
	if len(listtoobig) != 0:
		print("The given list of images exceeds the max size of eight images.")


# Function purpose is to sort images into the powerpoint with each slide being of modules with different conditions.
# Precondition: conditionlist is a full list of filenames with the same condition as strings.
# Postcondition: the function adds a slide and inserts all the images of the given condition into the slide.


def formatbycondition():
	#Creating this list just in case the number of images per slide will be more than 9 images.
		#POSSIBLE IMPROVEMENT AREA: Find some way to let the user know (maybe a pop up screen) when this happens.
	listtoobig = []

	#The user selected pre criteria from the GUI.
	condition = conditionentry.get()

	#Prompt the user to select the directory that contains all the possible image files they want to filter from.
	location = askdirectory(title = "Select location of image files")
	if(not location):
		sys.exit(0)

	#Change the current directory to the selected location.
	os.chdir(location)
	#Get a list of all files in the current directory that are jpeg images.
	filenames = glob.glob("*.jpg")

	#Find out how many photos have a module ID that matches the desired post condition.
	conditionlist = sortcondition(filenames, str(condition))

	#Select the right number of images to insert per slide.
	if len(conditionlist) == 0:
		print("There are no images that fit the given post condition")
	elif len(conditionlist) == 1:
		addoneimage(conditionlist)
	elif len(conditionlist) == 2:
		addtwoimage(conditionlist)
	elif len(conditionlist) == 3:
		addthreeimage(conditionlist)
	elif len(conditionlist) == 4:
		addfourimage(conditionlist)
	elif len(conditionlist) == 5:
		addfiveimage(conditionlist)
	elif len(conditionlist) == 6:
		addsiximage(conditionlist)
	elif len(conditionlist) == 7:
		addsevenimage(condiitonlist)
	elif len(conditionlist) == 8:
		addeightimage(conditionlist)
	else:
		listtoobig.append(file)

	#Currently the only indication that the user has tried putting more than 9 images per slide.
	if len(listtoobig) != 0:
		print("The given list of images exceeds the max size of eight images.")


# Precondition: a full list of image filenames is passed into the function.
# Postcondition: the function adds the 1 image from the list into the slide.


def addoneimage(modulelist):

	savename = savenameentry.get()

	pptx_location = askopenfilename(filetypes = [("Powerpoint Files", "*.pptx")], title = "Select Powerpoint Template File.")
	if not(pptx_location):
		sys.exit(0)

	prs = Presentation(pptx_location)

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

	width_full_standard_slide = Inches(10)
	width_full_widescreen_slide = Inches(13.333)
	height_full_slide = Inches(7.5)
	top_margin = Inches(1.5)
	side_margin = Inches(0.1)
	margin_in_between = Inches(0.5)

	top_1 = top_margin
	left_1 = side_margin
	width_1 = width_full_widescreen_slide-(2*left_1)
	height_1 = height_full_slide-(1.5*top_1)

	pictureslides.shapes.add_picture(modulelist[0], left_1, top_1, width_1, height_1)

	save_location = askdirectory(title = "Select location to save presentation.")
	os.chdir(save_location)
	prs.save(str(savename)+".pptx")


# Precondition: a full list of image filenames is passed into the function.
# Postcondition: the function adds the 2 images from the list into the slide.


def addtwoimage(modulelist):

	savename = savenameentry.get()

	pptx_location = askopenfilename(filetypes = [("Powerpoint Files", "*.pptx")], title = "Select Powerpoint Template File.")
	if not(pptx_location):
		sys.exit(0)

	prs = Presentation(pptx_location)

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
	
	width_full_standard_slide = Inches(10)
	width_full_widescreen_slide = Inches(13.333)
	height_full_slide = Inches(7.5)
	top_margin = Inches(1.5)
	side_margin = Inches(0.1)
	margin_in_between = Inches(0.5)

	top_2 = top_margin
	left_2_1 = side_margin
	left_2_2 = (width_full_widescreen_slide/2)+left_2_1
	width_2 = (width_full_widescreen_slide/2)-(2*left_2_1)
	height_2 = height_full_slide-(1.5*top_2)

	pictureslides.shapes.add_picture(modulelist[0], left_2_1, top_2, width_2, height_2)
	pictureslides.shapes.add_picture(modulelist[1], left_2_2, top_2, width_2, height_2)

	save_location = askdirectory(title = "Select location to save presentation.")
	os.chdir(save_location)
	prs.save(str(savename)+".pptx")


# Precondition: a full list of image filenames is passed into the function.
# Postcondition: the function adds the 3 images from the list into the slide.


def addthreeimage(modulelist):

	savename = savenameentry.get()

	pptx_location = askopenfilename(filetypes = [("Powerpoint Files", "*.pptx")], title = "Select Powerpoint Template File.")
	if not(pptx_location):
		sys.exit(0)

	prs = Presentation(pptx_location)

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
	
	width_full_standard_slide = Inches(10)
	width_full_widescreen_slide = Inches(13.333)
	height_full_slide = Inches(7.5)
	top_margin = Inches(1.5)
	side_margin = Inches(0.1)
	margin_in_between = Inches(0.5)

	top_3_1 = top_margin
	left_3_1 = side_margin
	top_3_2 = (height_full_slide/2)+(1.5*margin_in_between)
	left_3_2 = (width_full_widescreen_slide/2)+left_3_1
	width_3 = (width_full_widescreen_slide/2)-(2*left_3_1)
	height_3 = (height_full_slide/2)-top_3_1

	pictureslides.shapes.add_picture(modulelist[0], left_3_1, top_3_1, width_3, height_3)
	pictureslides.shapes.add_picture(modulelist[1], left_3_2, top_3_1, width_3, height_3)
	pictureslides.shapes.add_picture(modulelist[2], left_3_1, top_3_2, width_3, height_3)

	save_location = askdirectory(title = "Select location to save presentation.")
	os.chdir(save_location)
	prs.save(str(savename)+".pptx")


# Precondition: a full list of image filenames is passed into the function.
# Postcondition: the function adds the 4 images from the list into the slide.


def addfourimage(modulelist):

	savename = savenameentry.get()

	pptx_location = askopenfilename(filetypes = [("Powerpoint Files", "*.pptx")], title = "Select Powerpoint Template File.")
	if not(pptx_location):
		sys.exit(0)

	prs = Presentation(pptx_location)

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
	
	width_full_standard_slide = Inches(10)
	width_full_widescreen_slide = Inches(13.333)
	height_full_slide = Inches(7.5)
	top_margin = Inches(1.5)
	side_margin = Inches(0.1)
	margin_in_between = Inches(0.5)

	top_4_1 = top_margin
	left_4_1 = side_margin
	top_4_2 = (height_full_slide/2)+(1.5*margin_in_between)
	left_4_2 = (width_full_widescreen_slide/2)+left_4_1
	width_4 = (width_full_widescreen_slide/2)-(2*left_4_1)
	height_4 = (height_full_slide/2)-top_4_1

	pictureslides.shapes.add_picture(modulelist[0], left_4_1, top_4_1, width_4, height_4)
	pictureslides.shapes.add_picture(modulelist[1], left_4_2, top_4_1, width_4, height_4)
	pictureslides.shapes.add_picture(modulelist[2], left_4_1, top_4_2, width_4, height_4)
	pictureslides.shapes.add_picture(modulelist[3], left_4_2, top_4_2, width_4, height_4)

	save_location = askdirectory(title = "Select location to save presentation.")
	os.chdir(save_location)
	prs.save(str(savename)+".pptx")


# Precondition: a full list of image filenames is passed into the function.
# Postcondition: the function adds the 5 images from the list into the slide.


def addfiveimage(modulelist):

	savename = savenameentry.get()

	pptx_location = askopenfilename(filetypes = [("Powerpoint Files", "*.pptx")], title = "Select Powerpoint Template File.")
	if not(pptx_location):
		sys.exit(0)

	prs = Presentation(pptx_location)

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
	
	width_full_standard_slide = Inches(10)
	width_full_widescreen_slide = Inches(13.333)
	height_full_slide = Inches(7.5)
	top_margin = Inches(1.5)
	side_margin = Inches(0.1)
	margin_in_between = Inches(0.5)

	top_5_1 = top_margin
	left_5_1 = side_margin
	top_5_2 = (height_full_slide/2)+(1.5*margin_in_between)
	left_5_2 = (width_full_widescreen_slide/3)+left_5_1
	left_5_3 = (2*(width_full_widescreen_slide/3))+left_5_1
	width_5 = (width_full_widescreen_slide/3)-(2*left_5_1)
	height_5 = (height_full_slide/2)-top_5_1

	pictureslides.shapes.add_picture(modulelist[0], left_5_1, top_5_1, width_5, height_5)
	pictureslides.shapes.add_picture(modulelist[1], left_5_2, top_5_1, width_5, height_5)
	pictureslides.shapes.add_picture(modulelist[2], left_5_3, top_5_1, width_5, height_5)
	pictureslides.shapes.add_picture(modulelist[3], left_5_1, top_5_2, width_5, height_5)
	pictureslides.shapes.add_picture(modulelist[4], left_5_2, top_5_2, width_5, height_5)

	save_location = askdirectory(title = "Select location to save presentation.")
	os.chdir(save_location)
	prs.save(str(savename)+".pptx")


# Precondition: a full list of image filenames is passed into the function.
# Postcondition: the function adds the 6 images from the list into the slide.


def addsiximage(modulelist):

	savename = savenameentry.get()

	pptx_location = askopenfilename(filetypes = [("Powerpoint Files", "*.pptx")], title = "Select Powerpoint Template File.")
	if not(pptx_location):
		sys.exit(0)

	prs = Presentation(pptx_location)

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

	pictureslides.shapes.add_picture(modulelist[0], left_6_1, top_6_1, width_6, height_6)
	pictureslides.shapes.add_picture(modulelist[1], left_6_2, top_6_1, width_6, height_6)
	pictureslides.shapes.add_picture(modulelist[2], left_6_3, top_6_1, width_6, height_6)
	pictureslides.shapes.add_picture(modulelist[3], left_6_1, top_6_2, width_6, height_6)
	pictureslides.shapes.add_picture(modulelist[4], left_6_2, top_6_2, width_6, height_6)
	pictureslides.shapes.add_picture(modulelist[5], left_6_3, top_6_2, width_6, height_6)

	save_location = askdirectory(title = "Select location to save presentation.")
	os.chdir(save_location)
	prs.save(str(savename)+".pptx")


# Precondition: a full list of image filenames is passed into the function.
# Postcondition: the function adds the 7 images from the list into the slide.


def addsevenimage(modulelist):

	savename = savenameentry.get()

	pptx_location = askopenfilename(filetypes = [("Powerpoint Files", "*.pptx")], title = "Select Powerpoint Template File.")
	if not(pptx_location):
		sys.exit(0)

	prs = Presentation(pptx_location)

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
	
	width_full_standard_slide = Inches(10)
	width_full_widescreen_slide = Inches(13.333)
	height_full_slide = Inches(7.5)
	top_margin = Inches(1.5)
	side_margin = Inches(0.1)
	margin_in_between = Inches(0.5)

	top_7_1 = top_margin
	left_7_1 = side_margin
	top_7_2 = (height_full_slide/2)+(1.5*margin_in_between)
	left_7_2 = (width_full_widescreen_slide/4)+left_7_1
	left_7_3 = (2*(width_full_widescreen_slide/4))+left_7_1
	left_7_4 = (3*(width_full_widescreen_slide/4))+left_7_1
	width_7 = (width_full_widescreen_slide/4)-(2*left_7_1)
	height_7 = (height_full_slide/2)-top_7_1

	pictureslides.shapes.add_picture(modulelist[0], left_7_1, top_7_1, width_7, height_7)
	pictureslides.shapes.add_picture(modulelist[1], left_7_2, top_7_1, width_7, height_7)
	pictureslides.shapes.add_picture(modulelist[2], left_7_3, top_7_1, width_7, height_7)
	pictureslides.shapes.add_picture(modulelist[3], left_7_4, top_7_1, width_7, height_7)
	pictureslides.shapes.add_picture(modulelist[4], left_7_1, top_7_2, width_7, height_7)
	pictureslides.shapes.add_picture(modulelist[5], left_7_2, top_7_2, width_7, height_7)
	pictureslides.shapes.add_picture(modulelist[6], left_7_3, top_7_2, width_7, height_7)

	save_location = askdirectory(title = "Select location to save presentation.")
	os.chdir(save_location)
	prs.save(str(savename)+".pptx")
	


# Precondition: a full list of image filenames is passed into the function.
# Postcondition: the function adds the 8 images from the list into the slide.


def addeightimage(modulelist):

	savename = savenameentry.get()

	pptx_location = askopenfilename(filetypes = [("Powerpoint Files", "*.pptx")], title = "Select Powerpoint Template File.")
	if not(pptx_location):
		sys.exit(0)

	prs = Presentation(pptx_location)

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
	
	width_full_standard_slide = Inches(10)
	width_full_widescreen_slide = Inches(13.333)
	height_full_slide = Inches(7.5)
	top_margin = Inches(1.5)
	side_margin = Inches(0.1)
	margin_in_between = Inches(0.5)

	top_8_1 = top_margin
	left_8_1 = side_margin
	top_8_2 = (height_full_slide/2)+(1.5*margin_in_between)
	left_8_2 = (width_full_widescreen_slide/4)+left_8_1
	left_8_3 = (2*(width_full_widescreen_slide/4))+left_8_1
	left_8_4 = (3*(width_full_widescreen_slide/4))+left_8_1
	width_8 = (width_full_widescreen_slide/4)-(2*left_8_1)
	height_8 = (height_full_slide/2)-top_8_1

	pictureslides.shapes.add_picture(modulelist[0], left_8_1, top_8_1, width_8, height_8)
	pictureslides.shapes.add_picture(modulelist[1], left_8_2, top_8_1, width_8, height_8)
	pictureslides.shapes.add_picture(modulelist[2], left_8_3, top_8_1, width_8, height_8)
	pictureslides.shapes.add_picture(modulelist[3], left_8_4, top_8_1, width_8, height_8)
	pictureslides.shapes.add_picture(modulelist[4], left_8_1, top_8_2, width_8, height_8)
	pictureslides.shapes.add_picture(modulelist[5], left_8_2, top_8_2, width_8, height_8)
	pictureslides.shapes.add_picture(modulelist[6], left_8_3, top_8_2, width_8, height_8)
	pictureslides.shapes.add_picture(modulelist[7], left_8_4, top_8_2, width_8, height_8)

	save_location = askdirectory(title = "Select location to save presentation.")
	os.chdir(save_location)
	prs.save(str(savename)+".pptx")
	

# Precondition: function will take location of image files and then sort them into the separate lists.
# Postcondition: function will return a list of filenames all of the same criteria.


def getcriteriaimagefiles():
	location = askdirectory(title = "Select location of image files")
	if(not location):
		sys.exit(0)

	os.chdir(location)
	filenames = glob.glob("*.jpg")

	criterialist = sortcriteria(filenames, "EP002-P01")

	return criterialist


# Precondition: function will take location of image files and then sort them into the separate lists.
# Postcondition: function will return a list of filenames all of the same condition.


def getconditionimagefiles():
	location = askdirectory(title = "Select location of image files")
	if(not location):
		sys.exit(0)

	os.chdir(location)
	filenames = glob.glob("*.jpg")

	conditionlist = sortcondition(filenames, "DH500")

	return conditionlist


# Don't need this anymore {
# Program needs to be able to identify where the modules came from (since different places have different naming structure)
#	1. MST San Jose
#	2. MST Laguna
#	3. Waaree
# }

# Program will look at first 6 characters of filename and then it will know where the module came from

# Program needs to be able to identify the following parameters from the MST San Jose filename
#	1. Project number
#	2. Module size
#	3. Date manufactured
#	4. Coupon/module number
#	5. Coupon/module condition
#	6. Coupon image or EL image

# Program needs to be able to identify the following parameters from Laguna filename

# Program needs to be able to identify the folowing parameters from Waaree filename


# Write something here that extracts the #1-6 info requested from above, then creates separate lists to store files each of the same type
# Can use a dictionary to be more efficient




# Import all the images here

# Where all the files are placed
# location = askdirectory(title = "Select location of image files")
# if(not location):
# 	sys.exit(0)

# # # Location = "//SERVER-2/Shared Folders/Public/Engineering/GTSAN Bering Data/EL images Central"


# os.chdir(location)
# filenames = glob.glob("*.jpg")

#filenames = stem_list(filenames_ext)




# # Creating a list for each separate identifier


# criterialist = sortcriteria(filenames, "EP002-P01")
# datelist = sortdate(filenames, "2021", "March", "11")
# celltypelist = sortcelltype(filenames, "PERC G1 URE")
# encaplist = sortencap(filenames, "Cybrid POE Cybrid 0.2 POE")
# gridversionlist = sortgridversion(filenames, "Gen 2 Copper")
# mfglocationlist = sortmfglocation(filenames, "Laguna")
# familylist = sortfamily(filenames, "Specialty")
# colorlist = sortcolor(filenames, "Transparent")
# cellcutlist = sortcellcut(filenames, "Full")
# arrayconfiglist = sortarrayconfig(filenames, "12")
# conditionlist = sortcondition(filenames, "DH500")







# Part 2: Creating and formatting the powerpoint file

# Create and set up the powerpoint file
# Format, font, design, etc; aka import a premade template

# Open the powerpoint presentation template.


# pptx_location = askopenfilename(filetypes = [("Powerpoint Files", "*.pptx")], title = "Select Powerpoint Template File.")

# prs = Presentation(pptx_location)


# Define all the different types of slides in the powerpoint template.

# lyt_1 = prs.slide_layouts[0]
# lyt_2 = prs.slide_layouts[1]
# lyt_3 = prs.slide_layouts[2]
# lyt_4 = prs.slide_layouts[3]
# lyt_5 = prs.slide_layouts[4]
# lyt_6 = prs.slide_layouts[5]
# lyt_7 = prs.slide_layouts[6]
# lyt_8 = prs.slide_layouts[7]



# Create and set up the powerpoint file:
#	Ask the user which type of powerpoint they want by multiple choice
#	Code in several different types of powerpoint formats that we most commonly use


# How to size the images
# One option could be to look at the module size (which we got from the file name) and based on that decide how big the image should be?
# Image size is determined by module size and number of pictures


#	Different types of powerpoints:
#		1. Powerpoint that holds a lot of images
#			Some slides can have images while others can be of other stuff
#			Each slide contains one coupon, and the images in the slide are of the different conditions that the coupon went through
#			Some slides contain coupon images and some slides contain EL images


#		2. Powerpoint that holds a lot of images
#			Some slides can have images while others can be of other stuff
#			Each slide contains a different condition, and the images in the slide are of the different coupons that went through that condition
#			Some slides contain coupon images and some slides contain EL images
	

# if(ErrorFiles):
# 	showwarning(title='Warning!', message='The following file(s) could not be read: \n\n'+ErrorFiles)

# # Fill out the summary data on the summary sheet
# for p in range(1, len(headers)):
# 	clmn = chr(ord('A')+p)
# 	rnge = clmn+'5:'+clmn+str(sumrow)
# 	summary.write_formula(1,p,'=Min('+rnge+')',formlist[p])
# 	summary.write_formula(2,p,"=Average("+rnge+")", formlist[p])
# 	summary.write_formula(3,p,"=Max("+rnge+")",formlist[p])

# wb.close()


# Save the finished powerpoint file.


# save_location = askdirectory(title = "Select location to save presentation.")
# os.chdir(save_location)

# prs.save("test3.pptx")


# Part 3: Create the GUI


powerpointgui = tk.Tk()


# Create homepage screen.


powerpointcanvas = tk.Canvas(powerpointgui, width = 700, height = 600)
powerpointcanvas.pack()


# Create boxes to take user input.


criteriaentry = tk.Entry(powerpointgui)
powerpointcanvas.create_window(220, 100, window = criteriaentry)

criterialabel = tk.Label(text = "Pre Criteria")
powerpointcanvas.create_window(87.5, 100, window = criterialabel)

dateentry = tk.Entry(powerpointgui)
powerpointcanvas.create_window(220, 140, window = dateentry)

datelabel = tk.Label(text = "Date (YYMMDD)")
powerpointcanvas.create_window(87.5, 140, window = datelabel)

celltypeentry = tk.Entry(powerpointgui)
powerpointcanvas.create_window(220, 180, window = celltypeentry)

celltypelabel = tk.Label(text = "Cell Type")
powerpointcanvas.create_window(87.5, 180, window = celltypelabel)

encapentry = tk.Entry(powerpointgui)
powerpointcanvas.create_window(220, 220, window = encapentry)

encaplabel = tk.Label(text = "Encap")
powerpointcanvas.create_window(87.5, 220, window = encaplabel)

gridversionentry = tk.Entry(powerpointgui)
powerpointcanvas.create_window(220, 260, window = gridversionentry)

gridversionlabel = tk.Label(text = "Grid Version")
powerpointcanvas.create_window(87.5, 260, window = gridversionlabel)

mfglocationentry = tk.Entry(powerpointgui)
powerpointcanvas.create_window(220, 300, window = mfglocationentry)

mfglocationlabel = tk.Label(text = "Mfg Location")
powerpointcanvas.create_window(87.5, 300, window = mfglocationlabel)

familyentry = tk.Entry(powerpointgui)
powerpointcanvas.create_window(220, 340, window = familyentry)

familylabel = tk.Label(text = "Family")
powerpointcanvas.create_window(87.5, 340, window = familylabel)

colorentry = tk.Entry(powerpointgui)
powerpointcanvas.create_window(220, 380, window = colorentry)

colorlabel = tk.Label(text = "Color")
powerpointcanvas.create_window(87.5, 380, window = colorlabel)

cellcutentry = tk.Entry(powerpointgui)
powerpointcanvas.create_window(220, 420, window = cellcutentry)

cellcutlabel = tk.Label(text = "Cell Cut")
powerpointcanvas.create_window(87.5, 420, window = cellcutlabel)

arrayconfigentry = tk.Entry(powerpointgui)
powerpointcanvas.create_window(220, 460, window = arrayconfigentry)

arrayconfiglabel = tk.Label(text = "Array Configuration")
powerpointcanvas.create_window(87.5, 460, window = arrayconfiglabel)

conditionentry = tk.Entry(powerpointgui)
powerpointcanvas.create_window(220, 500, window = conditionentry)

conditionlabel = tk.Label(text = "Post Condition")
powerpointcanvas.create_window(87.5, 500, window = conditionlabel)

savenameentry = tk.Entry(powerpointgui)
powerpointcanvas.create_window(520, 100, window = savenameentry)

savenamelabel = tk.Label(text = "Powerpoint File Name")
powerpointcanvas.create_window(380, 100, window = savenamelabel)


# Create buttons to select powerpoint formatting.


formatbyconditionbutton = tk.Button(text = "Format Powerpoint by Post Condition", command = formatbycondition)
powerpointcanvas.create_window(450, 140, window = formatbyconditionbutton)

formatbycriteriabutton = tk.Button(text = "Format Powerpoint by Pre Criteria", command = formatbycriteria)
powerpointcanvas.create_window(450, 180, window = formatbycriteriabutton)

powerpointgui.mainloop()