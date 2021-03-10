# -*- coding: utf-8 -*-
"""
Created on Thu Jan 28 11:01:19 2021

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

def SimulationPurposesOnly():
	IO = generateIOMap()
	i = 0
	error = 0
	string = ""
	for row in data.iterrows():
		try:
			string += "MOV {0} Ind[{1}].Type ".format(IndType[row[1]["Type"]],row[1]["Indicator Index"])
			if i == 4:
				print(string)
				string = ""
				i = 0
			else:
				i+=1
		except:
			with open("IndicatorLog.txt","a") as errorlog:
				errorlog.write("Failed on {0}\n".format(row[1]["Indicator Index"]))
				error += 1
	print(string + "\n")
	if error>0:
		print("Errors occured while generating code. ({0})".format(error))

def SetTagValues():
	IO = generateIOMap()
	tags = ["HiHi","On","Mid","Off","LoLo","Value"]
	for row in data.iterrows():
		indicator = ""
		for tag in tags:
			if row[1][tag]!="":
				slot,point = row[1][tag].split("/")
				slot = IO[slot]
				indicator += "MOV {2} Ind[{0}].{1}_Slot MOV {3} Ind[{0}].{1}_Point ".format(row[1]["Indicator Index"], tag, slot, point)
				if tag == "Value":
					indicator += "OTL Ind[{0}].Analog_In ".format(row[1]["Indicator Index"])
		indicator += "MOV {1} Ind[{0}].Type ".format(row[1]["Indicator Index"],IndType[row[1]["Type"]])
		if indicator != "":
			print(indicator)
				
			

IndType = {"General/Level":1,
		   "Pressure/Vacuum":2,
		   "Speed/Switch":3}


dataFile = "//nw-srv22-file.ad.northwindts.com/Collaboration/Job Files/5631 Tuffy's Regrind Bins/Program Files/PLC Program/Regring PLC/Code Generator/Regrind_NewEquipment_DWG_Rev1-4.xlsx"
dataSheet = "Indicators"

data = pd.read_excel(dataFile,dataSheet).fillna("")


SetTagValues()