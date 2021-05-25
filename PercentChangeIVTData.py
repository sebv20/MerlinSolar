import glob
import os
import string
import os.path
import sys
import csv
import pyodbc
import pandas as pd
import xlsxwriter
import numpy as np
import matplotlib as plt
import seaborn as sns
import tkinter as tk
from tkinter.filedialog import *
from tkinter.messagebox import *
import Levenshtein




# Part 1: Copy needed columns from Access into Python and organize into an array.

def id_check(id_list,key):
	i = 0
	matches = {}
	#wanted_str = input("Enter your wanted character(s):")
	for match in id_list:
		if key in match:
			if match in matches:
				matches[match+"-SECOND-TEST"] = i
			else:
				matches[match] = i
		i+=1
	return matches

#def post_lam_find(wanted_id,id_list,matches,isc_list,voc_list,imp_list, vmp_list, pmp_list, ff_list, eff_list, rsh_list, rs_list):
#	post_lam_mod = "none"
#	for postlam in matches:
#		if "POST-LAM" in postlam:
#			if wanted_id in postlam:
#				post_lam_mod = postlam
#	#return post_lam_mod
#	i = 0
#	for wanted in id_list:
#		if wanted == post_lam_mod:
#			positions = i
#		i+=1
#
#	post_lam_dict = {"isc": isc_list[positions], "voc":voc_list[positions], "imp": imp_list[positions], "vmp": vmp_list[positions], "pmp": pmp_list[positions], "ff": ff_list[positions], "eff": eff_list[positions], "rsh": rsh_list[positions], "rs": rs_list[positions]}
#	return post_lam_dict


def post_lam(matches):

	condition = entry3.get()

	post_lam_list = []
	for post_lam in matches:
		if condition in post_lam:
			post_lam_list.append(post_lam)

	if len(post_lam_list) == 0:
		for post_lam4 in matches:
			if "POST-LAM" in post_lam4:
				post_lam_list.append(post_lam4)

	if len(post_lam_list) == 0:
		for post_lam2 in matches:
			if "POST-TRIM" in post_lam2:
				post_lam_list.append(post_lam2)

	if len(post_lam_list) == 0:
		for post_lam3 in matches:
			if "POST-JBOX" in post_lam3:
				post_lam_list.append(post_lam3)

	return post_lam_list




#function that will read through matches and find the correct modules from the batch 
def percent_change_helper(post_lams, wanted_id,matches,isc_list,voc_list,imp_list, vmp_list, pmp_list, ff_list, eff_list, rsh_list, rs_list):
	

	best_mod = post_lams[0]
	best_value = Levenshtein.distance(post_lams[0], wanted_id)
	for match in post_lams:
		num = Levenshtein.distance(match, wanted_id)
		if num < best_value:
			best_value = num
			best_mod = match
	for k,v in matches.items():
		if best_mod == k:
			post_lam_dict = {"isc": isc_list[v], "voc":voc_list[v], "imp": imp_list[v], "vmp": vmp_list[v], "pmp": pmp_list[v], "ff": ff_list[v], "eff": eff_list[v], "rsh": rsh_list[v], "rs": rs_list[v]}
	return post_lam_dict



def condition_split(mod_iden,matches):
	split_cond = []

	if mod_iden == "EP":
		for k,v in matches.items():
			temp = k.split('-',3)
			split_cond.append(temp[3])
		return split_cond

	else:
		for k,v in matches.items():
			temp = k.split('-',5)
			split_cond.append(temp[5])
		return split_cond

def id_split(mod_iden,matches):

	split_id = []

	if mod_iden == "EP":
		for k,v in matches.items():
			temp = k.split('-',3)
			tempdel = temp[3]
			temp.remove(tempdel)
			app = '-'.join(temp)
			split_id.append(app)
		return split_id
	else:
		for k,v in matches.items():
			temp = k.split('-',5)
			tempdel = temp[5]
			temp.remove(tempdel)
			app = '-'.join(temp)
			split_id.append(app)

		return split_id

#this function calculates percent change by first taking in lists with all values imported from access database in copyfromaccess()
#function then creates empty dictionaries to be filled by data 
# using indexing with numbers from the dictionary with matching ID (wanted_id), correct values are then filled in to their respective dictionaries 
def percent_change(mod_iden,matches,isc_list,voc_list,imp_list, vmp_list, pmp_list, ff_list, eff_list, rsh_list, rs_list):

	filename = entry1.get()
	post_lams = post_lam(matches)
	split_cond = condition_split(mod_iden,matches)
	split_ids = id_split(mod_iden, matches)

	percent_changeisc = {}
	percent_changevoc = {}
	percent_changeimp = {}
	percent_changevmp = {}
	percent_changepmp = {}
	percent_changeeff = {}
	percent_changeff = {}
	percent_changersh = {}
	percent_changers = {}

	wanted_isc = {}  
	wanted_voc = {}
	wanted_imp = {}
	wanted_vmp = {}
	wanted_pmp = {}
	wanted_eff = {}
	wanted_ff = {}
	wanted_rsh = {}
	wanted_rs = {}




	for k,v in matches.items():
		wanted_isc[k] = (isc_list[v])
		wanted_voc[k] = (voc_list[v])
		wanted_imp[k] = (imp_list[v])
		wanted_vmp[k] = (vmp_list[v])
		wanted_pmp[k] = (pmp_list[v])
		wanted_ff[k] = (ff_list[v])
		wanted_eff[k] = (eff_list[v])
		wanted_rsh[k] = (rsh_list[v])
		wanted_rs[k] = (rs_list[v])


	for k,v in matches.items():
		post_lam_dict = percent_change_helper(post_lams,k,matches,isc_list,voc_list,imp_list, vmp_list, pmp_list, ff_list, eff_list, rsh_list, rs_list)
		if isc_list[v] == None:
			percent_changeisc[k] = 0
		else:
			percent_changeisc[k] = ((isc_list[v] - post_lam_dict["isc"])/post_lam_dict["isc"]) *100
		if voc_list[v] == None:
			percent_changevoc[k] = 0
		else:
			percent_changevoc[k] = ((voc_list[v] - post_lam_dict["voc"])/post_lam_dict["voc"]) *100
		if imp_list[v] == None:
			percent_changeimp[k] = 0
		else:
			percent_changeimp[k] = ((imp_list[v] - post_lam_dict["imp"])/post_lam_dict["imp"]) *100
		if vmp_list[v] == None:
			percent_changevmp[k] = 0
		else:
			percent_changevmp[k] = ((vmp_list[v] - post_lam_dict["vmp"])/post_lam_dict["vmp"]) *100
		if pmp_list[v] == None:
			percent_changepmp[k] = 0
		else:
			percent_changepmp[k] = ((pmp_list[v] - post_lam_dict["pmp"])/post_lam_dict["pmp"]) *100
		if ff_list[v] == None:
			percent_changeff[k] = 0
		else:
			percent_changeff[k] = ((ff_list[v] - post_lam_dict["ff"])/post_lam_dict["ff"])*100
		if eff_list[v] == None:
			percent_changeeff[k] = 0
		else:
			percent_changeeff[k] = (eff_list[v] - post_lam_dict["eff"])
		if rsh_list[v] == None:
			percent_changersh[k] = 0
		else:
			percent_changersh[k] = ((rsh_list[v] - post_lam_dict["rsh"])/post_lam_dict["rsh"]) *100
		if rs_list[v] == None:
			percent_changers[k] = 0
		else:
			percent_changers[k] = ((rs_list[v] - post_lam_dict["rs"])/post_lam_dict["rs"]) *100

	#print(percent_changers)

	location = askdirectory(title = "Select location to export number of modules fit to.")
	os.chdir(location)

	workbook = xlsxwriter.Workbook(str(filename)+".xlsx") 

	#workbook = xlsxwriter.Workbook(excel_path)
	worksheet = workbook.add_worksheet()

	worksheet.write('A1', "Module ID and Condition")
	worksheet.write('B1', "Module ID")
	worksheet.write('C1', "Module Condition")
	worksheet.write('D1', 'ISC')
	worksheet.write('E1', 'VOC')
	worksheet.write('F1', 'IMP')
	worksheet.write('G1', 'VMP')
	worksheet.write('H1', 'PMP')
	worksheet.write('I1', 'FF')
	worksheet.write('J1', 'EFF')
	worksheet.write('K1', 'RSH')
	worksheet.write('L1', 'RS')	
	worksheet.write('M1', 'ISC % Change')
	worksheet.write('N1', 'VOC % Change')
	worksheet.write('O1', 'IMP % Change')
	worksheet.write('P1', 'VMP % Change')
	worksheet.write('Q1', 'PMP % Change')
	worksheet.write('R1', 'FF % Change')
	worksheet.write('S1', 'EFF % Change')
	worksheet.write('T1', 'RSH % Change')
	worksheet.write('U1', 'RS % Change')


	row=1 
	for	x in matches:
		worksheet.write(row,0, x)	
		row+=1		
	row=1 
	for	x in split_ids:
		worksheet.write(row,1, x)	
		row+=1	
	row=1 
	for	x in split_cond:
		worksheet.write(row,2, x)	
		row+=1			
	row=1 
	for k, v in wanted_isc.items():
		worksheet.write(row,3, v)
		row+=1
	row=1
	for k, v in wanted_voc.items():
		worksheet.write(row,4, v)
		row+=1
	row=1
	for k, v in wanted_imp.items():
		worksheet.write(row,5, v)
		row+=1
	row=1
	for k, v in wanted_vmp.items():
		worksheet.write(row,6, v)
		row+=1
	row=1
	for k, v in wanted_pmp.items():
		worksheet.write(row,7, v)
		row+=1
	row=1
	for k, v in wanted_ff.items():
		worksheet.write(row,8, v)
		row+=1
	row=1
	for k, v in wanted_eff.items():
		worksheet.write(row,9, v)
		row+=1
	row=1
	for k, v in wanted_rsh.items():
		worksheet.write(row,10, v)
		row+=1
	row=1
	for k, v in wanted_rs.items():
		worksheet.write(row,11, v)
		row+=1



	row=1
	for k, v in percent_changeisc.items():
		worksheet.write(row,12, v)
		row+=1
	row=1
	for k, v in percent_changevoc.items():
		worksheet.write(row,13, v)
		row+=1
	row=1
	for k, v in percent_changeimp.items():
		worksheet.write(row,14, v)
		row+=1
	row=1
	for k, v in percent_changevmp.items():
		worksheet.write(row,15, v)
		row+=1
	row=1
	for k, v in percent_changepmp.items():
		worksheet.write(row,16, v)
		row+=1
	row=1
	for k, v in percent_changeff.items():
		worksheet.write(row,17, v)
		row+=1
	row=1
	for k, v in percent_changeeff.items():
		worksheet.write(row,18, v)
		row+=1
	row=1
	for k, v in percent_changersh.items():
		worksheet.write(row,19, v)
		row+=1
	row=1
	for k, v in percent_changers.items():
		worksheet.write(row,20, v)
		row+=1
	workbook.close()
	print("done")



#	df = pd.DataFrame(data=percent_changeisc, index=[0])
#	df = (df.T)
#	print (df)
#	df.to_excel("data.xlsx")





def copyfromaccess():

	excel_path = entry1.get()
	module_id = entry2.get()

	accessfile = askopenfilename(filetypes = [("Access Files", "*.accdb")], title = 'Select location of Microsoft Access database.')
	#conn_str = (
		#r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};'
		#r'DBQ='+accessfile+';'
	#)

	print(accessfile)

	conn = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ='+accessfile+';')
	#conn = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=C:/Users/svargas/OneDrive - Merlin Solar Technologies, Inc/ModuleSummary.mdb;')
	conn.setdecoding(pyodbc.SQL_CHAR, encoding='utf8')
	conn.setdecoding(pyodbc.SQL_WCHAR, encoding='utf8')
	conn.setencoding(encoding='utf8')
	cursor = conn.cursor()
	#cursor.execute('Select * from Results')

	#saving each column of data to its own respective list 
	table_row = cursor.execute('Select * from Results')
	id_list = []
	for row in table_row:
		name = row[3]
		id_list.append(name)
	#print(id_list)


	table_row = cursor.execute('Select * from Results')
	isc_list = []
	for row in table_row:
		isc_value = row[5]
		isc_list.append(isc_value)
	#print(len(isc_list))
	#print(isc_list)

	table_row = cursor.execute('Select * from Results')
	voc_list = []
	for row in table_row:
		voc_value = row[6]
		voc_list.append(voc_value)
	#print(len(voc_list))
	#print(voc_list)

	table_row = cursor.execute('Select * from Results')
	imp_list = []
	for row in table_row:
		imp_value = row[7]
		imp_list.append(imp_value)
	#print(len(imp_list))
	#print(imp_list)


	table_row = cursor.execute('Select * from Results')
	vmp_list = []
	for row in table_row:
		vmp_value = row[8]
		vmp_list.append(vmp_value)
	#print(len(vmp_list))
	#print(vmp_list)

	table_row = cursor.execute('Select * from Results')
	pmp_list = []
	for row in table_row:
		pmp_value = row[9]
		pmp_list.append(pmp_value)
	#print(len(pmp_list))
	#print(pmp_list)

	table_row = cursor.execute('Select * from Results')
	ff_list = []
	for row in table_row:
		ff_value = row[10]
		ff_list.append(ff_value)
	#print(len(ff_list))
	#print(ff_list)

	table_row = cursor.execute('Select * from Results')
	eff_list = []
	for row in table_row:
		eff_value = row[12]
		eff_list.append(eff_value)
	#print(len(eff_list))
	#print(eff_list)

	table_row = cursor.execute('Select * from Results')
	rsh_list = []
	for row in table_row:
		rsh_value = row[13]
		rsh_list.append(rsh_value)
	#print(len(rsh_list))
	#print(rsh_list)

	table_row = cursor.execute('Select * from Results')
	rs_list = []
	for row in table_row:
		rs_value = row[14]
		rs_list.append(rs_value)
	#print(len(rs_list))
	#print(rs_list)


	matches = {}
	matches = id_check(id_list, module_id)
	#print(matches)

	#calling post lam function to find the module with post lam in the name and saving the result 
	#post_lam = post_lam_find(id_list,matches,isc_list,voc_list,imp_list, vmp_list, pmp_list, ff_list, eff_list, rsh_list, rs_list)
	#print(post_lam)
	mod_iden = next(iter(matches))

	if mod_iden[0:2] == "EP":
		mod_iden = "EP"


	percent_change(mod_iden,matches,isc_list,voc_list,imp_list, vmp_list, pmp_list, ff_list, eff_list, rsh_list, rs_list)

	label4 = tk.Label(lengthwidth, text = "The percent change data has been exported to the Excel spreadsheet selected earlier.")
	canvas1.create_window(350, 340, window = label4)






# Part 2: Calculate percent change for all IVT data for all conditions from POSTLAM condition.


lengthwidth = tk.Tk()

canvas1 = tk.Canvas(lengthwidth, width = 700, height = 500)
canvas1.pack()

entry1 = tk.Entry(lengthwidth)
canvas1.create_window(450, 140, window = entry1)

label1 = tk.Label(text = "Name of Excel File to be Created")
canvas1.create_window(240, 140, window = label1)

entry2 = tk.Entry(lengthwidth)
canvas1.create_window(450, 180, window = entry2)

label2 = tk.Label(text = "Module ID to be Searched")
canvas1.create_window(240, 180, window = label2)

entry3 = tk.Entry(lengthwidth)
canvas1.create_window(450, 220, window = entry3)

label3 = tk.Label(text = "Module Condition to be used as Baseline")
canvas1.create_window(240, 220, window = label3)



# Part 2: Return a list of packing factors for all modules for both orientations (no mixing of anything).


button1 = tk.Button(text = "Enter", command = copyfromaccess)
canvas1.create_window(350, 260, window = button1)

lengthwidth.mainloop()



#a = copyfromaccess()

#return a

#i=0
#	wanted_list = []
#	wanted_key = 'P2941'
#	while i< len(name_list):
#		x=0
#		checker = True
#		while x<5:
#			if wanted_list[x] != wanted_key[x]:
#				checker = False
#	if checker == True
#		wanted_list.append(name_list[i])
#	print(wanted_list)

#	i=0
#	while i< len(name_list):
#		if id_check(name_list[i], 'P2941') == True:
#			wanted_list.append(name_list[i])
#	print(wanted_list)


#C:/Users/svargas/OneDrive - Merlin Solar Technologies, Inc/Data1.xlsx
#P2828-M01-2X1-200820-05
