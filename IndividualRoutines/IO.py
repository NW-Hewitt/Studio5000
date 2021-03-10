# -*- coding: utf-8 -*-
"""
Created on Tue Aug  4 15:36:44 2020

@author: Joe.Hewitt
"""


import pandas as pd


RockwellCatalog = './Data/Catalog.xlsx'
cardDB = pd.read_excel(RockwellCatalog,'Basic').fillna('')
#cardDB.set_index('Model')


def front(x,n=0):
	y = str(x)
	#print(len(y),y,n)
	if len(y)>=n:
		return(y)
	while len(y)<n:
		y = '0'+ y
		#print(y)
	return(y)

class Mod:
	
	
	def __init__(self, card, rack, slot, QLslot):
		#print(f'Looking up {card} in catalog...')
		cardInfo = cardDB.loc[cardDB['Model']==card]
		#print(cardInfo)
		#print('Card found.  Extracting information...')
		self.catalog = (cardInfo.values[0][0]).split('-')[0]
		self.rack = rack
		self.slot = slot
		self.ql = QLslot
		self.type=cardInfo.values[0][1]
		self.IO=cardInfo.values[0][2]
		self.InTag = cardInfo.values[0][4]
		self.OutTag = cardInfo.values[0][6]
		
		if self.IO == 'both':
			self.ICount,self.OCount =  cardInfo.values[0][3].split('/')
			self.ICount = int(self.ICount)
			self.OCount = int(self.OCount)
		elif self.IO == 'input':
			self.ICount = cardInfo.values[0][3]
			self.OCount = 0
		elif self.IO == 'output':
			self.ICount = 0
			self.OCount = cardInfo.values[0][3]
			
		if not cardInfo.values[0][5]:
			self.ICount=1
		if not cardInfo.values[0][7]:
			self.OCount=1
		self.digits = cardInfo.values[0][8]
		#print(self.digits)
		#print(f'New {card} module ready for use.')
		
	def genIO(self):
		tmpString = ''
		#print('Generating QL IO code')
		if self.catalog == '5069':
			if self.InTag!='':
				tmpString += self.createCode5069(self.InTag, self.ICount, [self.rack,self.slot,0,self.ql])
			if self.OutTag!='':
				tmpString += self.createCode5069(self.OutTag, self.OCount, [self.rack,self.slot,0,self.ql])
			
		else:
			if self.InTag!='':
				tmpString += self.createCode(self.InTag, self.ICount, [self.rack,self.slot,0,self.ql])
			if self.OutTag!='':
				tmpString += self.createCode(self.OutTag, self.OCount, [self.rack,self.slot,0,self.ql])
		
		self.Studio5000Code = tmpString
			
		
	def createCode(self, tagString, iterLimit, variableArray):
		tmpString = ''
		fullString = ''
		for i in range(iterLimit):
			variableArray[2] = front(i,n=self.digits)
			tmpString += tagString.format(*variableArray)
			if len(tmpString) > 130:
				fullString += tmpString
				tmpString = '\n'
		if len(tmpString)>3:
			fullString+=tmpString
		return(fullString)
	
	def createCode5069(self, tagString, iterLimit, variableArray):
		tmpString = ''
		fullString = ''
		if (('Ainput' in tagString) or ('Aoutput' in tagString)):
			for i in range(iterLimit):
				variableArray[2] = front(i,n=self.digits)
				tmpString += tagString.format(*variableArray)
				if (len(tmpString) > 130):
					fullString += tmpString
					tmpString = '\n'
		else:
			tmpString += 'BST '
			for i in range(iterLimit):
				if i != 0:
					tmpString += 'NXB '
				variableArray[2] = front(i,n=self.digits)
				tmpString += tagString.format(*variableArray)
			tmpString += 'BND '
				
				
		if len(tmpString)>3:
			fullString+=tmpString
		return(fullString)
			


