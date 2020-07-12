#saml excel data 
import os 
import glob
import pandas as pd 
import re #bruges til at replace ting i 
import numpy as np 
import warnings
warnings.filterwarnings('ignore') # fordi noget excel kode giver nogle warnings
import tkinter as tk # til userinterface 
from tkinter import ttk
from tkinter import messagebox
import xml.etree.ElementTree as ET 



def find_files(path):
	os.chdir(path)						#change til path
	files = glob.glob('*.xml') 			#virker kun i working directory. har nu liste med files
	print('')
	print('')
	print('Number of files: {}'.format(len(files)))
	print('Number of files: {}'.format(len(files)))
	print('Number of files: {}'.format(len(files)))
	print('')
	print('')
	return files 


def get_tapse(fp): 
	rawdata = ET.parse(fp)
	data = rawdata.getroot()

	#retrieve CPR 
	cpr = data.find('.//PatientId').text #retrieve patient ID node 

	#retrieve measurements node 
	measurements = data.find('.//Series')

	#retrieve in a list (returns a list!)
	param = measurements.findall('Parameter')

	#intilalise tapse list
	tapse_values = []

	for x in param: 												#loop throgh parameters in study
		if x.find('ParameterName').text == 'TAPSE':					#if parameter name is TAPSE
			if x.find('DisplayUnit').text == 'cm':					#if displayunit is cm (to fix the weird tapse cm/s and so on errors in Livs data)
				tapse_values.append(x.find('DisplayValue').text)	#append value to tapse vales list

	#create panda series 
	tapse = pd.Series(tapse_values)

	#convert from string to numeric
	tapse = pd.to_numeric(tapse)

	#take mean value of tapse values 
	mean_tapse = tapse.mean()

	#create pandas series for dataframe creation
	cpr = pd.Series(cpr, name ='cpr')
	tapse = pd.Series(mean_tapse, name = 'tapse' ) 

	#stack columns 
	Df = pd.concat([cpr,tapse], axis = 1) #stack along axis = 1 / stack columns (axis = 0 for row stack)

	#return tapse for pt 
	return Df 


################################################################################################################################
################################################################################################################################
#get the tapse values ##########################get the tapse values ##########################get the tapse values ##########################get the tapse values ##########################get the tapse values #########################
#get the tapse values ####################get the tapse values ##########################get the tapse values ##########################get the tapse values ##########################get the tapse values ##########################get the tapse values #########################
################################################################################################################################
################################################################################################################################

path = '/Users/danielmodin/Desktop/data'

#find files 
files = find_files(path)

#initializer
i = 0 

#run through XML files 
for file in files: 
	print(files)
	if i == 0:
		Df = get_tapse(file)
	if i > 0:
		temp = get_tapse(file)
		Df = pd.concat([Df,temp], axis = 0)
	i = i+1 


#drop pts with NaN (we dont need patients who do not have tapse)
Df = Df.dropna() 

#fix cpr nummer problems 
Df.cpr = Df.cpr.replace('-','',regex=True)
Df.cpr = Df.cpr.replace('#','',regex=True)
Df.cpr = Df.cpr.replace('migration','',regex=True)



#gem data 
print(' ')
print('Writing data to XL file')
print('Writing data to XL file')
print('Writing data to XL file')
Df.to_excel('tapse.xlsx', engine = 'xlsxwriter', index = False)
print('Done writing XL')
print('Done writing XL')
print('Done writing XL')






