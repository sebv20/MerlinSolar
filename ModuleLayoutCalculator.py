import glob
import os
import string
import os.path
import sys
import csv
import xlsxwriter
import pandas as pd
import numpy as np
import matplotlib as plt
import seaborn as sns
import tkinter as tk
from tkinter.filedialog import *
from tkinter.messagebox import *
from datetime import date
from pathlib import Path


# Define program related functions here.


# Precondition: function takes the length and width of the roof inputted by the user.
# Postcondition: function returns a list containing the areawise packing factor for each type of module.


def areapackingcalculator(length, width):
	areapackingfactor = []
	lengthmodulesfit = []
	widthmodulesfit = []
	modulestoobig = []
	errorfiles = []

	modulelist = [
	
	"RWS-P0100BSRO ", 
	"RWS-P0320BQRO ", 
	"RWS-P0360BFRO ", 
	"RWS-P0440BHLO ", 
	"RWS-P0480BFLO ", 
	"RWS-P0560BQLO ", 
	"RWS-P0720BFRP ",
	"RWS-P0840WFRO ",
	"TBS-P0230WHRO ",
	"TBS-P0410WHLO ",
	"TBS-P0480WFLO ",
	"TBS-P0480WHRG ",
	"TBS-P0480WHRO ",
	"TBS-P0720WFRO ",
	"TBS-P0720WFRP ",
	"RWS-P0890WFRO ",
	"RWS-P0360BHSO ",
	"RWS-P0350BHRO ",
	"RWS-P0240BFLO ",
	"RWS-P0480BFSO ",
	"TBS-P0420WHSO ",
	"TBS-P0830WHRO ",
	"TBS-P0240WFSO ",
	"TBS-P0180WHLO ",
	"TBS-P0360WFLO ",
	"TBS-P0360WFSO ",
	"RWS-P0360BFDO "
	
	]

	lengthlist = [

	142,
	746,
	1554,
	1871,
	2033,
	1245,
	1963,
	2285,
	1020,
	1167,
	2033,
	1067,
	1067,
	1963,
	1963,
	2456,
	975,
	910,
	2038,
	1306,
	857,
	1182,
	1062,
	1070,
	3003,
	1060,
	2030

	]

	widthlist = [

	182,
	373,
	689,
	376,
	683,
	373,
	1051,
	1061,
	356,
	540,
	683,
	703,
	685,
	1051,
	1051,
	1020,
	606,
	606,
	372,
	1068,
	783,
	1013,
	669,
	285,
	372,
	987,
	536

	]

	arealist = np.multiply(lengthlist, widthlist)

	# arealist = lengthlist*widthlist
	roofarea = float(length)*float(width)

	for lengthvalue in lengthlist:
		if float(lengthvalue) <= float(length):
			lengthnummodulesfit = int(float(length)/float(lengthvalue))
			lengthmodulesfit.append(lengthnummodulesfit)
		else:
			lengthnummodulesfit = int(float(length)/float(lengthvalue))
			lengthmodulesfit.append(lengthnummodulesfit)


	for widthvalue in widthlist:
		if float(widthvalue) <= float(width):
			widthnummodulesfit = int(float(width)/float(widthvalue))
			widthmodulesfit.append(widthnummodulesfit)
		else:
			widthnummodulesfit = int(float(width)/float(widthvalue))
			widthmodulesfit.append(widthnummodulesfit)

	areamodulesfit = np.multiply(lengthmodulesfit, widthmodulesfit)
	totalmodulearealist = np.multiply(np.array(arealist, dtype = np.int64), np.array(areamodulesfit), dtype = np.int64)

	# print(lengthmodulesfit)
	# print(widthmodulesfit)
	# print(arealist)
	# print(areamodulesfit)
	# print(totalmodulearealist)
	# print(roofarea)

	for areavalue in totalmodulearealist:
		if float(areavalue) <= roofarea:
			areapackfactor = areavalue/roofarea
			areapackingfactor.append(str(areapackfactor))
		elif float(areavalue) >= roofarea:
			modulestoobig.append(areavalue)
		else:
			errorfiles.append(areavalue)

	areapackingfactorwithmodules = [i + j for i, j in zip(modulelist, areapackingfactor)]

	splitareapackingfactor = []

	for item in areapackingfactorwithmodules:
		splitresult = item.split()
		itemlist = [splitresult[0], splitresult[1]]
		splitareapackingfactor.append(itemlist)

	if len(modulestoobig) != 0:
		print("These modules are too big for the given roof: ", modulestoobig)

	if len(errorfiles) != 0:
		print("An unknown lengthwise error occured: ", errorfiles)

	return splitareapackingfactor


# Precondition: function takes the length and width of the roof inputted by the user.
# Postcondition: function returns a list containing the number of modules that can fit on the given roof.


def areamodulefitcalculator(length, width):
	lengthmodulesfit = []
	widthmodulesfit = []
	modulestoobig = []
	errorfiles = []

	modulelist = [
	
	"RWS-P0100BSRO ", 
	"RWS-P0320BQRO ", 
	"RWS-P0360BFRO ", 
	"RWS-P0440BHLO ", 
	"RWS-P0480BFLO ", 
	"RWS-P0560BQLO ", 
	"RWS-P0720BFRP ",
	"RWS-P0840WFRO ",
	"TBS-P0230WHRO ",
	"TBS-P0410WHLO ",
	"TBS-P0480WFLO ",
	"TBS-P0480WHRG ",
	"TBS-P0480WHRO ",
	"TBS-P0720WFRO ",
	"TBS-P0720WFRP ",
	"RWS-P0890WFRO ",
	"RWS-P0360BHSO ",
	"RWS-P0350BHRO ",
	"RWS-P0240BFLO ",
	"RWS-P0480BFSO ",
	"TBS-P0420WHSO ",
	"TBS-P0830WHRO ",
	"TBS-P0240WFSO ",
	"TBS-P0180WHLO ",
	"TBS-P0360WFLO ",
	"TBS-P0360WFSO ",
	"RWS-P0360BFDO "
	
	]

	lengthlist = [

	142,
	746,
	1554,
	1871,
	2033,
	1245,
	1963,
	2285,
	1020,
	1167,
	2033,
	1067,
	1067,
	1963,
	1963,
	2456,
	975,
	910,
	2038,
	1306,
	857,
	1182,
	1062,
	1070,
	3003,
	1060,
	2030

	]

	widthlist = [

	182,
	373,
	689,
	376,
	683,
	373,
	1051,
	1061,
	356,
	540,
	683,
	703,
	685,
	1051,
	1051,
	1020,
	606,
	606,
	372,
	1068,
	783,
	1013,
	669,
	285,
	372,
	987,
	536

	]

	arealist = np.multiply(lengthlist, widthlist)

	# arealist = lengthlist*widthlist
	roofarea = float(length)*float(width)

	for lengthvalue in lengthlist:
		if float(lengthvalue) <= float(length):
			lengthnummodulesfit = int(float(length)/float(lengthvalue))
			lengthmodulesfit.append(lengthnummodulesfit)
		else:
			lengthnummodulesfit = int(float(length)/float(lengthvalue))
			lengthmodulesfit.append(lengthnummodulesfit)

	for widthvalue in widthlist:
		if float(widthvalue) <= float(width):
			widthnummodulesfit = int(float(width)/float(widthvalue))
			widthmodulesfit.append(widthnummodulesfit)
		else:
			widthnummodulesfit = int(float(width)/float(widthvalue))
			widthmodulesfit.append(widthnummodulesfit)

	areamodulesfit = np.multiply(lengthmodulesfit, widthmodulesfit)
	stringareamodulesfit = [str(i) for i in areamodulesfit]

	areamodulesfitwithmodules = [i + j for i, j in zip(modulelist, stringareamodulesfit)]

	splitnummodulesfit = []

	for item in areamodulesfitwithmodules:
		splitresult = item.split()
		itemlist = [splitresult[0], splitresult[1]]
		splitnummodulesfit.append(itemlist)

	if len(modulestoobig) != 0:
		print("These modules are too big for the given roof: ", modulestoobig)

	if len(errorfiles) != 0:
		print("An unknown lengthwise error occured: ", errorfiles)

	return splitnummodulesfit


# Define GUI related functions here.


def getpackingfactor():
	rooflength = entry1.get()
	roofwidth = entry2.get()

	packingfactor = areapackingcalculator(rooflength, roofwidth)

	location = askdirectory(title = "Select location to export packing factors to.")
	os.chdir(location)

	filename = entry3.get()

	with xlsxwriter.Workbook(str(filename)+".xlsx") as workbook:
		worksheet = workbook.add_worksheet()

		worksheet.write(0, 0, "Module Serial Number")
		worksheet.write(0, 1, "Module Packing Factor")
		worksheet.write(0, 3, "Roof Dimensions")
		worksheet.write(1, 3, "Length (mm)")
		worksheet.write(2, 3, "Width (mm)")
		worksheet.write(1, 4, float(rooflength))
		worksheet.write(2, 4, float(roofwidth))

		row = 1
		col = 0

		for module, packfactor in packingfactor:
			worksheet.write(row, col, module)
			worksheet.write(row, col + 1, float(packfactor))
			row += 1

		#workbook.close()

	label5 = tk.Label(lengthwidth, text = "The packing factor of each module is exported to an Excel spreadsheet located in the selected folder.")
	canvas1.create_window(350, 340, window = label5)

def getnummodules():
	rooflength = entry1.get()
	roofwidth = entry2.get()

	modulesfit = areamodulefitcalculator(rooflength, roofwidth)

	location = askdirectory(title = "Select location to export number of modules fit to.")
	os.chdir(location)

	filename = entry4.get()

	with xlsxwriter.Workbook(str(filename)+".xlsx") as workbook:
		worksheet = workbook.add_worksheet()

		worksheet.write(0, 0, "Module Serial Number")
		worksheet.write(0, 1, "Number of Modules that Fit on Roof")
		worksheet.write(0, 3, "Roof Dimensions")
		worksheet.write(1, 3, "Length (mm)")
		worksheet.write(2, 3, "Width (mm)")
		worksheet.write(1, 4, float(rooflength))
		worksheet.write(2, 4, float(roofwidth))

		row = 1
		col = 0

		for module, numfit in modulesfit:
			worksheet.write(row, col, module)
			worksheet.write(row, col + 1, float(numfit))
			row += 1

		#workbook.close()

	label6 = tk.Label(lengthwidth, text = "The number of modules that fit is exported to an Excel spreadheet located in the selected folder.")
	canvas1.create_window(350, 380, window = label6)


# Part 1: Prompt user to input the length and width of the rectangular roof in meters.


lengthwidth = tk.Tk()

canvas1 = tk.Canvas(lengthwidth, width = 700, height = 500)
canvas1.pack()

entry1 = tk.Entry(lengthwidth)
canvas1.create_window(350, 100, window = entry1)

label1 = tk.Label(text = " Roof Length (mm)")
canvas1.create_window(220, 100, window = label1)

entry2 = tk.Entry(lengthwidth)
canvas1.create_window(350, 140, window = entry2)

label2 = tk.Label(text = "Roof Width (mm)")
canvas1.create_window(220, 140, window = label2)

entry3 = tk.Entry(lengthwidth)
canvas1.create_window(450, 180, window = entry3)

label3 = tk.Label(text = "Packing Factor Spreadsheet Filename")
canvas1.create_window(240, 180, window = label3)

entry4 = tk.Entry(lengthwidth)
canvas1.create_window(450, 220, window = entry4)

label4 = tk.Label(text = "Number of Modules Fit Spreadsheet Filename")
canvas1.create_window(240, 220, window = label4)


# Part 2: Return a list of packing factors for all modules for both orientations (no mixing of anything).


button1 = tk.Button(text = "Packing Factor of All Modules", command = getpackingfactor)
canvas1.create_window(350, 260, window = button1)

button2 = tk.Button(text = "Number of Modules that Fit", command = getnummodules)
canvas1.create_window(350, 300, window = button2)

lengthwidth.mainloop()
