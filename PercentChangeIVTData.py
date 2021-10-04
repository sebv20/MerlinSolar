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
from datetime import datetime, timedelta



#goes through list of all id's to find ones that match desired id, also labels multiple matches as "second test"
def id_check(id_list,key):
	i = 0
	matches = {}
	for match in id_list:
		if key in match:
			if match in matches:
				matches[match+"-SECOND-TEST"] = i
			else:
				matches[match] = i
		i+=1
	return matches

def date_check(date_list,id_list,key):
	i = 0
	matches = {}
	#print(key)
	#print(date_list[0].strftime("%m/%d/%Y"))
	for match in date_list:
		if str(match.strftime("%m/%d/%Y")) == key:
			if id_list[i] in matches:
				matches[id_list[i]+"-SECOND-TEST"] = i
			else:
				matches[id_list[i]] = i
		i+=1
	matches_final = matches
	mod_id = id_split("EP", matches)
	key_list = []
	value_list = []
	q=0
	for key, value in matches.items():
		add = post_lam_helper(mod_id[q], id_list)
		if add != 10000000:
			#matches_final[id_list[add]] = add 
			key_list.append(id_list[add])
			value_list.append(add)
		q+=1
	g=0
	for x in key_list:
		matches_final[x] = value_list[g]
		g+=1

	#print(matches_final)
	return matches_final



def test_check(id_list,date_list, dateEntry, key):

	newMatches = {}
	matches = date_check(date_list,id_list,dateEntry)
	print(matches)
	for match in matches.keys():
		if key in match:
			if match in newMatches:
				newMatches[match+"SECOND-TEST"] = matches.get(match)
				print("TWO")
			else:
				newMatches[match] = matches.get(match)
	matches_final = newMatches
	mod_id = id_split("EP", newMatches)
	key_list = []
	value_list = []
	q=0
	for key, value in newMatches.items():
		add = post_lam_helper(mod_id[q], id_list)
		if add != 10000000:
			#matches_final[id_list[add]] = add 
			key_list.append(id_list[add])
			value_list.append(add)
		q+=1
	g=0
	for x in key_list:
		matches_final[x] = value_list[g]
		g+=1

	return matches_final



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
		if "POST-LAM" in post_lam:
			post_lam_list.append(post_lam)
		if "POSTLAM" in post_lam:
			post_lam_list.append(post_lam)

	if len(post_lam_list) == 0:
		for post_lam2 in matches:
			if "POST-TRIM" in post_lam2:
				post_lam_list.append(post_lam2)

	if len(post_lam_list) == 0:
		for post_lam3 in matches:
			if "POST-JBOX" in post_lam3:
				post_lam_list.append(post_lam3)

	return post_lam_list


#this helps find a post lam using modules id's chosen by date 
def post_lam_helper(module, id_list):
	condition =  entry3.get()
	backup = "POST-LAM"
	backup2 = "POSTLAM"
	if condition == "":
		condition = "POSTLAM"
	i=0
	#print(module)
	for x in id_list:
		if module in x:
			#print(x)
			if condition in x:
			#	print("found all")
				return i 
			if backup in x:
			#	print("found all")
				return i
			if backup2 in x:
			#	print("found all")
				return i
		i+=1
	return 10000000




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
		try:
			for k,v in matches.items():
				temp = k.split('-',3)
				split_cond.append(temp[3])
			return split_cond
		except:
			for k,v in matches.items():
				temp = k.split('-',1)
				split_cond.append(temp[1])
			return split_cond			

	else:
		for k,v in matches.items():
			temp = k.split('-',5)
			split_cond.append(temp[5])
		return split_cond


def id_split(mod_iden,matches):

	split_id = []

	if mod_iden == "EP":
		try:
			for k,v in matches.items():
				temp = k.split('-',3)
				tempdel = temp[3]
				temp.remove(tempdel)
				app = '-'.join(temp)
				split_id.append(app)
			return split_id
			#print(split_id)
		except:
			for k,v in matches.items():
				temp = k.split('-',1)
				tempdel = temp[1]
				temp.remove(tempdel)
				app = '-'.join(temp)
				split_id.append(app)
			return split_id		
			#print(split_id)
	else:
		for k,v in matches.items():
			temp = k.split('-',5)
			tempdel = temp[5]
			temp.remove(tempdel)
			app = '-'.join(temp)
			split_id.append(app)

		return split_id
		#print(split_id)

def barcode_helper(row, column, match_string):

	year = {'A': '2020', 'B': '2021', 'C': '2022', 'D': '2023', 'E': '2024', 'F': '2025', 'G': '2026', 'H': '2027', 'I': '2028', 'J': '2029', 'K': '2030', 'L': '2031'}
	month = {'A': 'January', 'B': 'February', 'C': 'March', 'D': 'April', 'E': 'May', 'F': 'June', 'G': 'July', 'H': 'August', 'I': 'September', 'J': 'October', 'K': 'November', 'L': 'December'}
	day = {'1': '1', '2': '2', '3': '3', '4': '4', '5': '5', '6': '6', '7': '7', '8': '8', '9': '9', 'A': '10', 'B': '11', 'C':'12', 'D': '13', 'E':'14', 'F': '15', 'G': '16', 'H': '17', 'J': '18','K': '19','L': '20', 'M': '21', 'N': '22','P': '23','R': '24', 'S': '25','T': '26', 'U': '27','V': '28', 'W': '29', 'X': '30', 'Y': '31'}
	cell_type = {'_': 'Parent', 'A': 'Mono M2 EEPV', 'B': 'PERC M2 EEPV', 'C': 'PERC G1 URE', 'D': 'HJT M2 + HEVEL', 'E': 'HJT M2 + Kaneka', 'F': 'Perc M2 URE', 'G': 'Provisional 1', 'H': 'Provisional 1', 'Z': 'Engineering 1'}
	encap = {"_": "Parent", "A": "Mistui EVA", 'B': 'Cybrid POE', 'C': 'Mitsui POE', 'D': 'First EVA STR EVA 0.2', 'E': "First EVA 0.45mm", 'G': 'Cybrid POE 0.45 STR EVA 0.2', 'F': 'Provisional 1', 'Z': 'Engineering 1'}
	grid_version = {'1': 'POR', '2': 'Gen 2 Black', '3': "Gen 2 Copper", '4': 'Provisional 1', '5': 'Provisional 1', '0': 'Engineering 1'}
	mfg_location = {'1': 'San Jose', '2': 'Waaree', '3': 'Laguna'}
	family = {'F': "FX Transporation", 'R': 'FX Roofing', 'B': 'Back Rail Glass Roofing', 'M': 'FX Marine', 'N': 'Nishati', 'H': 'Honeycomb Board', 'G': 'GX', 'X': 'FX Portables/Folding', 'S': 'Specialty'}
	color = {'B': 'Black', 'W': 'White', 'T': 'Transparent', 'C': 'Camo', 'M': 'Multicolor', 'P': 'Provisional 1', 'Q': 'Provisional 1', 'X': 'Engineering 1'}
	cell_cut  = {'F': 'Full', 'H': 'Half', 'Q': 'Quarter', 'E': 'Eighth', 'S': 'Sixteenth'}
	array_config = [['10 SC Config (Trucklite)','F', 'AA'],['2x5 TC Config', 'T', 'AA' ], ['3x(6x6) QC XP Config', 'P', 'AA'], ['4x14 QC Config', 'F', 'AB'],['4x8 QC Config', 'F', 'AC'],['8x4 QC Config', 'F', 'AD'],['12x6 QC Config', 'F', 'AE'],['14x6 QC Config', 'F', 'AF'],['3x(4x5) HC Multifold (Nishati)', 'N', 'AA'],
	['8x6 HC Config', 'F', 'AB'],['4x12 HC Config', 'F', 'AC'], ['4x11 HC Config', 'F', 'AD'], ['3x(2x7) HC BXD Config', 'P', 'AE'], ['6x7-1 HC Config', 'F', 'AF'],['6x6 HC horizontal', 'F', 'AG'],['3x(4x3) HC XP Config', 'P', 'AH'],['5x7 HC Config', 'F', 'AI'],['2x(4x4) HC Multifold (Nishati)', 'N', 'AJ'], ['4x(2x3) HC MiniXP Config', 'N', 'AK'],
	['5x4 HC Config – King', 'F', 'AL'],['4x(2x2) HC MiniXP Config', 'P', 'AM'],['3X6 HC Config', 'F', 'AN'],['4x6-1 HC Config', 'F', 'AO'],['4x6 HC Config', 'F', 'AP'],['6x4 HC Config', 'F', 'AQ'],['6x7 HC Config', 'F', 'AR'],['6x14 HC Config 84 HC Back JB', 'F', 'AS'],['6x14-1 HC Config 83 HC Top JB', 'F', 'AT'],['9x5-3 42HC Config', 'F', 'AU'],
	['6x12 FC BR Config', 'R', 'AA'],['6x12 FC S Config', 'F', 'AB'],['6x12 FC P,Corner Jbox, Turnt', 'F', 'AC'],['6x12 FC P,Corner Jbox, Standard', 'F', 'AD'],['8x6 FC Config', 'F', 'AE'],['4x12 FC Config', 'F', 'AF'],['3x14-1 FC Config', 'F', 'AG'],['3x(4x3) FC XP Config', 'F', 'AH'],['3x12 FC Config', 'F', 'AI'],['2x18 FC Config', 'F','AJ'],['2x18 FC Config','R', 'AJ'],
	['6x6 FC Config', 'F','AK'],['6x6 FC Config','R', 'AK'],['4x9 FC Config', 'F', 'AL'],['4x8 FC Config', 'F', 'AM'],['2x(4x4) FC Multifold (Nishati)', 'N', 'AN'],['2x15 FC Config', 'F', 'AO'],['3x(4x2) FC Trifold Nishati', 'N', 'AP'],['2x12 FC Config','R', 'AQ'],['2x12 FC Config', 'F','AQ'],['3x8 FC Config', 'F', 'AR'],['4x6 FC Config', 'F', 'AS'],['3x7 FC Config', 'F', 'AT'],
	['2x9 FC Config', 'F', 'AU'],['2x8 FC Config', 'F', 'AV'],['4x4 FC Config', 'F', 'AW'],['4x3 FC Config', 'F', 'AX'],['2x6 FC Config', 'F', 'AY'],['2X3 FC Config', 'F', 'AZ'],['12x8 FC TinyWatts', 'F', 'BA'],['8x6 FC Navistar', 'F', 'BB'],['14x6 FC Config', 'F', 'BC'],['1x1 FC Config','F','BD'],['1x2 FC Config','F','BE'],['2x2 FC Config','F','BF'],
	['2x3 FC Config','F','BG'],['3x3 FC Config', 'F','BH'],['4x4 FC Config','F','BI']]
	#NEEDS TO BE EXCEL-ED
	#CHECK BARCODE FOR FAMILY MATCHING VALUES
	index = 0
	counter =0
	start = 0
	for x in match_string:
		if x == '-':
			if counter == 1:
				start = index+2
			counter+=1
		else:
			index+=1

	barcode = match_string[start:]
	output = []

	if len(barcode) >10:
		for k,v in year.items():
			if k == barcode[0]:
				output.append(v)

		for k,v in month.items():
			if k == barcode[1]:
				output.append(v)

		for k,v in day.items():
			if k == barcode[2]:
				output.append(v)

		for k,v in cell_type.items():
			if k == barcode[3]:
				output.append(v)

		for k,v in encap.items():
			if k == barcode[4]:
				output.append(v)

		for k,v in grid_version.items():
			if k == barcode[5]:
				output.append(v)

		for k,v in mfg_location.items():
			if k == barcode[6]:
				output.append(v)

		for k,v in family.items():
			if k == barcode[7]:
				output.append(v)

		for k,v in color.items():
			if k == barcode[8]:
				output.append(v)

		for k,v in cell_cut.items():
			if k == barcode[9]:
				output.append(v)

		for x in array_config:
			# print(x[2])
			# print(barcode[10:11])
			if x[2] == barcode[10:12]:
				# print(barcode[10:12])
				# if x[1] == barcode[7]:
				output.append(x[0])
				#print(x[0])				

		return output
	else:
		return output



# This function essentially performs the same way as percent_change, but is optimized for the "test" analyzation method 
#def percent_change_test
#	filename = entr


#this function calculates percent change by first taking in lists with all values imported from access database in copyfromaccess()
#function then creates empty dictionaries to be filled by data 
# using indexing with numbers from the dictionary with matching ID (wanted_id), correct values are then filled in to their respective dictionaries 
def percent_change(mod_iden,matches,isc_list,voc_list,imp_list, vmp_list, pmp_list, ff_list, eff_list, rsh_list, rs_list):

	filename = entry1.get()
	post_lams = post_lam(matches)
	#print(len(post_lams))
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
		if k[0] == "E":
			post_lam_dict = percent_change_helper(post_lams,k,matches,isc_list,voc_list,imp_list, vmp_list, pmp_list, ff_list, eff_list, rsh_list, rs_list)
			if isc_list[v] == None:
				percent_changeisc[k] = 0
			elif post_lam_dict["isc"] == None:
				post_lam_dict["isc"] == 0
			else:
				percent_changeisc[k] = ((isc_list[v] - post_lam_dict["isc"])/post_lam_dict["isc"]) *100
			if voc_list[v] == None:
				percent_changevoc[k] = 0
			elif post_lam_dict["voc"] == None:
				post_lam_dict["voc"] == 0
			else:
				percent_changevoc[k] = ((voc_list[v] - post_lam_dict["voc"])/post_lam_dict["voc"]) *100
			if imp_list[v] == None:
				percent_changeimp[k] = 0
			elif post_lam_dict["imp"] == None:
				post_lam_dict["imp"] == 0
			else:
				percent_changeimp[k] = ((imp_list[v] - post_lam_dict["imp"])/post_lam_dict["imp"]) *100
			if vmp_list[v] == None:
				percent_changevmp[k] = 0
			elif post_lam_dict["vmp"] == None:
				post_lam_dict["vmp"] == 0
			else:
				percent_changevmp[k] = ((vmp_list[v] - post_lam_dict["vmp"])/post_lam_dict["vmp"]) *100
			if pmp_list[v] == None:
				percent_changepmp[k] = 0
			elif post_lam_dict["pmp"] == None:
				post_lam_dict["pmp"] == 0
			else:
				percent_changepmp[k] = ((pmp_list[v] - post_lam_dict["pmp"])/post_lam_dict["pmp"]) *100
			if ff_list[v] == None:
				percent_changeff[k] = 0
			elif post_lam_dict["ff"] == None:
				post_lam_dict["ff"] == 0
			else:
				percent_changeff[k] = ((ff_list[v] - post_lam_dict["ff"])/post_lam_dict["ff"])*100
			if eff_list[v] == None:
				percent_changeeff[k] = 0
			elif post_lam_dict["eff"] == None:
				post_lam_dict["eff"] == 0
			else:
				percent_changeeff[k] = (eff_list[v] - post_lam_dict["eff"])
			if rsh_list[v] == None:
				percent_changersh[k] = 0
			elif post_lam_dict["rsh"] == None:
				post_lam_dict["rsh"] == 0
			else:
				percent_changersh[k] = ((rsh_list[v] - post_lam_dict["rsh"])/post_lam_dict["rsh"]) *100
			if rs_list[v] == None:
				percent_changers[k] = 0
			elif post_lam_dict["rs"] == None:
				post_lam_dict["rs"] == 0
			else:
				percent_changers[k] = ((rs_list[v] - post_lam_dict["rs"])/post_lam_dict["rs"]) *100

	#print(percent_changers)

	location = askdirectory(title = "Select location to export excel file to.")
	os.chdir(location)

	workbook = xlsxwriter.Workbook(str(filename)+".xlsx") 

	#workbook = xlsxwriter.Workbook(excel_path)
	worksheet = workbook.add_worksheet()


	worksheet.write('A1', 'Year')
	worksheet.write('B1', 'Month')
	worksheet.write('C1', 'Day')
	worksheet.write('D1', 'Cell Type')
	worksheet.write('E1', 'Encap')
	worksheet.write('F1', 'Grid Version')
	worksheet.write('G1', 'MFG Location')
	worksheet.write('H1', 'Family')
	worksheet.write('I1', "Color")
	worksheet.write('J1', 'Cell Cut')
	worksheet.write('K1', 'Array Config')

	worksheet.write('L1', "Module ID and Condition")
	worksheet.write('M1', "Module ID")
	worksheet.write('N1', "Module Condition")
	worksheet.write('O1', 'ISC')
	worksheet.write('P1', 'VOC')
	worksheet.write('Q1', 'IMP')
	worksheet.write('R1', 'VMP')
	worksheet.write('S1', 'PMP')
	worksheet.write('T1', 'FF')
	worksheet.write('U1', 'EFF')
	worksheet.write('V1', 'RSH')
	worksheet.write('W1', 'RS')	
	worksheet.write('X1', 'ISC % Change')
	worksheet.write('Y1', 'VOC % Change')
	worksheet.write('Z1', 'IMP % Change')
	worksheet.write('AA1', 'VMP % Change')
	worksheet.write('AB1', 'PMP % Change')
	worksheet.write('AC1', 'FF % Change')
	worksheet.write('AD1', 'EFF % Change')
	worksheet.write('AE1', 'RSH % Change')
	worksheet.write('AF1', 'RS % Change')


	row=1 
	column = 0
	#print(len(split_ids))
	for x in split_ids: 
		column =0
		barcode_list = barcode_helper(row, column, x)
		for y in barcode_list:
			worksheet.write(row, column, y)
			column+=1
		row+=1
	row=1
	column=11
	for	x in matches:
		worksheet.write(row,column, x)	
		row+=1	
	column+=1	
	row=1 
	for	x in split_ids:
		worksheet.write(row,column, x)	
		row+=1	
	column+=1
	row=1 
	for	x in split_cond:
		worksheet.write(row,column, x)	
		row+=1	
	column+=1		
	row=1 
	for k, v in wanted_isc.items():
		worksheet.write(row,column, v)
		row+=1
	column+=1
	row=1
	for k, v in wanted_voc.items():
		worksheet.write(row,column, v)
		row+=1
	column+=1
	row=1
	for k, v in wanted_imp.items():
		worksheet.write(row,column, v)
		row+=1
	column+=1
	row=1
	for k, v in wanted_vmp.items():
		worksheet.write(row,column, v)
		row+=1
	column+=1
	row=1
	for k, v in wanted_pmp.items():
		worksheet.write(row,column, v)
		row+=1
	column+=1
	row=1
	for k, v in wanted_ff.items():
		worksheet.write(row,column, v)
		row+=1
	column+=1
	row=1
	for k, v in wanted_eff.items():
		worksheet.write(row,column, v)
		row+=1
	column+=1
	row=1
	for k, v in wanted_rsh.items():
		worksheet.write(row,column, v)
		row+=1
	column+=1
	row=1
	for k, v in wanted_rs.items():
		worksheet.write(row,column, v)
		row+=1
	column+=1
	row=1
	for k, v in percent_changeisc.items():
		worksheet.write(row,column, v)
		row+=1
	row=1
	column+=1
	for k, v in percent_changevoc.items():
		worksheet.write(row,column,v)
		row+=1
	row=1
	column+=1
	for k, v in percent_changeimp.items():
		worksheet.write(row,column, v)
		row+=1
	row=1
	column+=1
	for k, v in percent_changevmp.items():
		worksheet.write(row,column, v)
		row+=1
	row=1
	column+=1
	for k, v in percent_changepmp.items():
		worksheet.write(row,column, v)
		row+=1
	row=1
	column+=1
	for k, v in percent_changeff.items():
		worksheet.write(row,column, v)
		row+=1
	row=1
	column+=1
	for k, v in percent_changeeff.items():
		worksheet.write(row,column, v)
		row+=1
	row=1
	column+=1
	for k, v in percent_changersh.items():
		worksheet.write(row,column, v)
		row+=1
	row=1
	column+=1
	for k, v in percent_changers.items():
		worksheet.write(row,column, v)
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
	dateEntry = entry4.get()
	testEntry = entry5.get()

	analyzeMethod = dropdown()

	accessfile = askopenfilename(filetypes = [("Access Files", "*.accdb")], title = 'Select location of Microsoft Access database.')
	#conn_str = (
		#r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};'
		#r'DBQ='+accessfile+';'
	#)

	#print(accessfile)

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
	date_list = []
	for row in table_row:
		date = row[4]
		date_list.append(date)
	#print(date_list)


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
		eff_value = row[11]
		eff_list.append(eff_value)
	#print(len(eff_list))
	#print(eff_list)

	table_row = cursor.execute('Select * from Results')
	rsh_list = []
	for row in table_row:
		rsh_value = row[12]
		rsh_list.append(rsh_value)
	#print(len(rsh_list))
	#print(rsh_list)

	table_row = cursor.execute('Select * from Results')
	rs_list = []
	for row in table_row:
		rs_value = row[13]
		rs_list.append(rs_value)
	#print(len(rs_list))
	#print(rs_list)


	# if statements to check which is the desired analyzation method 
	if analyzeMethod == "Project":
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
		canvas1.create_window(350, 420, window = label4)

	if analyzeMethod == "Date":
		matches = {}
		matches = date_check(date_list, id_list, dateEntry)
		#print(matches)
		mod_iden = next(iter(matches))

		if mod_iden[0:2] == "EP":
			mod_iden = "EP"

		#print(mod_iden)
		percent_change(mod_iden,matches,isc_list,voc_list,imp_list, vmp_list, pmp_list, ff_list, eff_list, rsh_list, rs_list)
		label4 = tk.Label(lengthwidth, text = "The percent change data has been exported to the Excel spreadsheet selected earlier.")
		canvas1.create_window(350, 420, window = label4)


	if analyzeMethod == "Test":
		matches = {}
		matches = test_check(id_list, date_list, dateEntry, testEntry)
		#print(matches)

		mod_iden = next(iter(matches))

		if mod_iden[0:2] == "EP":
			mod_iden = "EP"

		percent_change(mod_iden,matches,isc_list,voc_list,imp_list, vmp_list, pmp_list, ff_list, eff_list, rsh_list, rs_list)
		label4 = tk.Label(lengthwidth, text = "The percent change data has been exported to the Excel spreadsheet selected earlier.")
		canvas1.create_window(350, 420, window = label4)




#User Text Box Entries

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

entry4 = tk.Entry(lengthwidth)
canvas1.create_window(450,260,window = entry4)

label4 = tk.Label(text = "Date (Ex; '09/20/2021')")
canvas1.create_window(240,260, window = label4)

entry5 = tk.Entry(lengthwidth)
canvas1.create_window(450,300, window = entry5)

label5 = tk.Label(text = "Test (Ex; 'TC', 'HAST')")
canvas1.create_window(240,300, window = label5)

tkvar = tk.StringVar(lengthwidth)
choices = {'Project', 'Date', 'Test'}
tkvar.set('Project')

popUpMenu = tk.OptionMenu(lengthwidth, tkvar, *choices)
choiceLabel = tk.Label(text = 'Analyze By:')
canvas1.create_window(240, 340, window=choiceLabel)
canvas1.create_window(450,340, window= popUpMenu)

def dropdown(*args):
	return tkvar.get()

tkvar.trace('w', dropdown)



# Part 2: Return a list of packing factors for all modules for both orientations (no mixing of anything).


button1 = tk.Button(text = "Enter", command = copyfromaccess)
canvas1.create_window(350, 380, window = button1)

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

 
