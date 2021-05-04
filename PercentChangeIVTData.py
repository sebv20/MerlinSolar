import glob
import os
import string
import os.path
import sys
import csv
import xlsxwriter
import pyodbc
import pandas as pd
import numpy as np
import matplotlib as plt
import seaborn as sns
import tkinter as tk
from tkinter.filedialog import *
from tkinter.messagebox import *


# Part 1: Copy needed columns from Access into Python and organize into an array.


def copyfromaccess():
	accessfile = askopenfilename(Filetypes = [("Access Files", "*.accdb")], title = 'Select location of Microsoft Access database.')

	conn_str = (
		r'DRIVER={Microsoft Access Driver (*.mdb, *accdb)};'
		r'DBQ=accessfile;'
	)

	conn = pyodbc.connect(conn_str)
	cursor = conn.cursor()
	cursor.execute('select * from ModuleSummary')

	for row in cursor.fetchall():
		print(row)

# Part 2: Calculate percent change for all IVT data for all conditions from POSTLAM condition.


#def calculatepercentchange():


a = copyfromaccess()
print(a)