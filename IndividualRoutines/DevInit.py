# -*- coding: utf-8 -*-
"""
Created on Thu Jan 28 08:18:36 2021

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
			string += "MOV {0} Dev[{1}].Type ".format(DevType[row[1]["Type"]],row[1]["Device Index"])
			if i == 4:
				print(string)
				string = ""
				i = 0
			else:
				i+=1
		except:
			with open("DevioceLog.txt","a") as errorlog:
				errorlog.write("Failed on {0}\n".format(row[1]["Device Index"]))
				error+=1
	print(string + "\n")
	if error>0:
		print("Errors occured while generating code. ({0})".format(error))
	
def SetTagValues():
	index=0
	IO = generateIOMap()
	
	tags = ["Energize","Energize2","Auxiliary","Auxiliary2","Cmd_Rate","Act_Rate","Fault_Remote","RemReset","Disconnect","Safety_ILK","Man_ILK"]
	for row in data.iterrows():
		device = ""
		if row[1]["Type"] != "VFD Motor":
			for tag in tags:
				if row[1][tag]!="":
					slot,point = row[1][tag].split("/")
					slot = IO[slot]
					device += "MOV {2} dev[{0}].{1}_Slot MOV {3} dev[{0}].{1}_Point ".format(row[1]["Device Index"], tag, slot, point)
					if tag == "Value":
						device += "OTL dev[{0}].Analog_In ".format(row[1]["Device Index"])
			device += "MOV {1} dev[{0}].Type ".format(row[1]["Device Index"],DevType[row[1]["Type"]])
		else:
			slot = IO[row[1]["PID"]]
			device = "BST MOV 1 Dev[{0}].Energize_Point MOV 1 Dev[{0}].Auxiliary_Point MOV 0 Dev[{0}].Cmd_Rate_Point MOV 0 Dev[{0}].Act_Rate_Point MOV 7 Dev[{0}].Fault_Remote_Point MOV 3 Dev[{0}].RemReset_Point ".format(row[1]["Device Index"])
			device += "NXB "
			for vfdtag in ["Energize","Auxiliary","Cmd_Rate","Act_Rate","Fault_Remote","RemReset"]:
				device += "MOV {1} Dev[{0}].{2}_Slot ".format(row[1]["Device Index"],slot,vfdtag)
			device += "BND "
		if device != "":
			print(device)
			index +=1
			if index == 10:
				print(" ")
				index = 0
			
			
DevType = {"On/Off Motor":1,
		   "VFD Motor":2,
		   "Single-Solenoid Valve":3,
		   "Dual-Solenoid Valve":4,
		   "Unproofed Device":5,
		   "Reversing On/Off Motor":6,
		   "Reversing VFD Motor":7,
		   "Control Valve":8,
		   "Slow/Fast Proofed":9,
		   "Slow/Fast Unproofed":10,
		   "Duty Cycle":11}

'''
5631
Batching
//nw-srv22-file.ad.northwindts.com/Collaboration/Job Files/5631 Tuffy's Regrind Bins/Program Files/PLC Program/Batching PLC/Code Generator/Batching_NewEquipment_DWG_Rev1-4.xlsx
Regrind
//nw-srv22-file.ad.northwindts.com/Collaboration/Job Files/5631 Tuffy's Regrind Bins/Program Files/PLC Program/Regring PLC/Code Generator/Regrind_NewEquipment_DWG_Rev1-4.xlsx
'''

dataFile = "//nw-srv22-file.ad.northwindts.com/Collaboration/Job Files/5631 Tuffy's Regrind Bins/Program Files/PLC Program/Regring PLC/Code Generator/Regrind_NewEquipment_DWG_Rev1-4.xlsx"
dataSheet = "Devices"

data = pd.read_excel(dataFile,dataSheet).fillna("")


SetTagValues()
	
	
