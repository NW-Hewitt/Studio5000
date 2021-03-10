# -*- coding: utf-8 -*-
"""
Created on Fri Feb  5 14:34:07 2021

@author: Joe.Hewitt
"""

import pandas as pd
from datetime import datetime

inputfile = "//nw-srv22-file.ad.northwindts.com/Collaboration/Job Files/5631 Tuffy's Regrind Bins/Program Files/PLC Program/Batching PLC/Code Generator/Copy of Batching_NewEquipment_DWG_Rev1_4-PRGM7_2.xlsx"

stamp = datetime.now().strftime("%Y%m%d_%H%M%S_")

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
		   "Duty Cycle":11,
		   "":0}

IndType = {"General/Level":1,
		   "Pressure/Vacuum":2,
		   "Speed/Switch":3}



def generateIOMap():
	IODict = {}
	iodata = pd.read_excel(inputfile,"IO Cards").fillna("")
	for row in iodata.iterrows():
		name = ""
		if row[1]["DWG Name"]!="":
			name = row[1]["DWG Name"]
		else:
			if row[1]["Name"]!="":
				name = row[1]["Name"]
		if name != "":
			IODict[name] = row[1]["PLC Index"]
	return(IODict)

def DeviceTags():
	index=0
	data = pd.read_excel(inputfile,"Devices").fillna("")
	tags = ["Energize","Energize2","Auxiliary","Auxiliary2","Cmd_Rate","Act_Rate","Fault_Remote","RemReset","Disconnect","Safety_ILK","Man_ILK"]
	with open(f"./StudioCode/{stamp}Device.txt","w+") as output:
		for row in data.iterrows():
			#print(row,'\n')
			device = ""
			if row[1]["Type"] != "VFD Motor":
				for tag in tags:
					if row[1][tag]!="":
						slot,point = row[1][tag].split("/")
						try:
							slot = IO[slot]
						except:
							print(slot)
						device += "MOV {2} dev[{0}].{1}_Slot MOV {3} dev[{0}].{1}_Point ".format(row[1]["Device Index"], tag, slot, point)
				device += "MOV {1} dev[{0}].Type ".format(row[1]["Device Index"],DevType[row[1]["Type"]])
			else:
				try:
					slot = IO[row[1]["PID"]]
				except:
					print(row[1]["Device Index"],row[1]["PID"])
				device = "BST MOV 1 Dev[{0}].Energize_Point MOV 1 Dev[{0}].Auxiliary_Point MOV 0 Dev[{0}].Cmd_Rate_Point MOV 0 Dev[{0}].Act_Rate_Point MOV 7 Dev[{0}].Fault_Remote_Point MOV 3 Dev[{0}].RemReset_Point ".format(row[1]["Device Index"])
				device += "NXB "
				for vfdtag in ["Energize","Auxiliary","Cmd_Rate","Act_Rate","Fault_Remote","RemReset"]:
					device += "MOV {1} Dev[{0}].{2}_Slot ".format(row[1]["Device Index"],slot,vfdtag)
				device += "BND "
			if device != "":
				output.write(device+"\n")
				index +=1
				if index == 10:
					output.write("\n")
					index = 0
				
def IndicatorTags():
	data = pd.read_excel(inputfile,"Indicators").fillna("")
	tags = ["HiHi","On","Mid","Off","LoLo","Value"]
	with open(f"./StudioCode/{stamp}Indicators.txt","w+") as output:
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
				output.write(indicator)
			
def InBitTags():
	data = pd.read_excel(inputfile,"UIB").fillna("")
	string = ""
	n = 0
	with open(f"./StudioCode/{stamp}UIB.txt","w+") as output:
		for row in data.iterrows():
			bit = row[1]["UIB Index"]
			slot,point = row[1]["Input"].split("/")
			slot = IO[slot]
			string += f"MOV {slot} InBit_Slot[{bit}] MOV {point} Inbit_Point[{bit}] "
			n+=1
			if n==5:
				n=0
				output.write(string)
				string = ""
		output.write(string)

def OutBitTags():
	data = pd.read_excel(inputfile,"UOB").fillna("")
	string = ""
	n = 0
	with open(f"./StudioCode/{stamp}UOB.txt","w+") as output:
		for row in data.iterrows():
			bit = row[1]["UOB Index"]
			slot,point = row[1]["Output"].split("/")
			slot = IO[slot]
			string += f"MOV {slot} OutBit_Slot[{bit}] MOV {point} Outbit_Point[{bit}] "
			n+=1
			if n==5:
				n=0
				output.write(string)
				string = ""
		output.write(string)

def DriveTags():
	data = pd.read_excel(inputfile,"Drives").fillna("")
	with open(f"./StudioCode/{stamp}Drives.txt","w+") as output:
		for row in data.iterrows():
			if row[1]["Model"] == "Powerflex 525-EENET":
				output.write('PF_525_IO IO_525 {0}:I {0}:O Dev[{2}] DInputs[{1}] DOutputs[{1}] AOutputs[{1},0] AInputs[{1},0] \n'.format(row[1]["Drive Name"],row[1]["Slot"],row[1]["Device Index"]))
	
def __main__():
	print("Generating IO Map...")
	IO = generateIOMap()
	print("Creating Device file...")
	DeviceTags()
	print("Creating Indicator file...")
	IndicatorTags()
	print("Creating UserInputBit file...")
	InBitTags()
	print("Creating UserOutputBit file...")
	OutBitTags()
	print("Creating Drive IO file...")
	DriveTags()


__main__()



