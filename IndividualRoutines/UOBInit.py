# -*- coding: utf-8 -*-
"""
Created on Wed Feb  3 12:37:21 2021

@author: Joe.Hewitt
"""

import pandas as pd

def generateIOMap():
	IODict = {}
	iodata = pd.read_excel(dataFile,"IO Cards").fillna("")
	for row in iodata.iterrows():
		name = ""
		if row[1]["DWG Name"]!="":
			name = row[1]["DWG Name"]
			
		else:
			if row[1]["Name"]!="":
				name = row[1]["Name"]
		if name != "":
			IODict.update({name:row[1]["PLC Index"]})
	return(IODict)

def SetTagValues():
	IO = generateIOMap()
	string = ""
	n = 0
	for row in data.iterrows():
		bit = row[1]["UOB Index"]
		slot,point = row[1]["Output"].split("/")
		slot = IO[slot]
		string += f"MOV {slot} OutBit_Slot[{bit}] MOV {point} Outbit_Point[{bit}] "
		n+=1
		if n==5:
			n=0
			print(string)
			string = ""
	print(string)
	
	
	
dataFile = "//nw-srv22-file.ad.northwindts.com/Collaboration/Job Files/5631 Tuffy's Regrind Bins/Program Files/PLC Program/Regring PLC/Code Generator/Regrind_NewEquipment_DWG_Rev1-4.xlsx"
dataSheet = "UOB"

data = pd.read_excel(dataFile,dataSheet).fillna("")


SetTagValues()