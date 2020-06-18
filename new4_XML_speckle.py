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


#globals 
df_xml = None 
df_xl = None 
df2_xl = None 



#Håndter fil indlæsning
def excel_find_files(path):
	#path = input('Enter path: ')		#Get path 
	os.chdir(path)						#change til path
	files = glob.glob('*.xls') 			#virker kun i working directory. har nu liste med files
	print('')
	print('')
	print('Number of files: {}'.format(len(files)))
	return files 



def excel_process(filepath):
	########################################Load excel nr. 1 #####################################
	########################################Load excel nr. 1 #####################################
	df = pd.read_excel(filepath, sheet_name='GlobalView_LAX')														
	#Fix cpr numre og column navne 
	cpr = df['ID']
	cpr = cpr.replace('#','',regex=True) 							#Panda har en replace( arguments .. ) funktion for dataframes. regex = True kræves for string
	cpr = cpr.replace('-','',regex=True)
	cpr = cpr.replace('migration','',regex=True)
	df['ID'] = cpr 
	del cpr 
	names = df.columns.str.strip(' ').str.replace(' ','_').str.replace('%','').str.replace('(','').str.replace(')','').str.replace('/','').str.replace('migration','').str.replace('migrationmigration','')
	df.columns = names 
	#Melt data 
	df2 = pd.melt(df, id_vars=['ID','View'], value_name='value', var_name='variable')
	df2['variable'] = df2['variable'].astype(str)
	df2['View'] = df2['View'].astype(str) 
	df2['variable'] = df2['View']+'_'+df2['variable'] #samler view+variabel 
	del df2['View']
	df2['value'] = pd.to_numeric(df2.value, errors = 'coerce') 	#convert to numeric, og lav non-numeric strings til missing
	df3 = df2.groupby(['ID','variable'], as_index = False).agg({'value':'mean'})  # Midle dobbelt målinger!!        
	df3 = df3.pivot(index = 'ID', columns = 'variable', values = 'value' )
	#Færdig longitudinelt global
	df_global_lax = df3
	#print(df_global_lax)
	########################################Load excel nr. 2 #####################################
	########################################Load excel nr. 2 #####################################
	df = pd.read_excel(filepath, sheet_name='SegmVal_LAX')
	cpr = df['ID']
	cpr = cpr.replace('#','',regex=True) 							#Panda har en replace( arguments .. ) funktion for dataframes. regex = True kræves for string
	cpr = cpr.replace('-','',regex=True)
	cpr = cpr.replace('migration','',regex=True)
	df['ID'] = cpr 
	del cpr 
	names = df.columns.str.strip(' ').str.replace(' ','_').str.replace('%','').str.replace('(','').str.replace(')','').str.replace('/','').str.replace('migration','').str.replace('migrationmigration','')
	df.columns = names
	#melt data
	df2 = pd.melt(df, id_vars=['ID','Segment'], value_name='value', var_name='variable')
	df2['variable'] = df2['variable'].astype(str)
	df2['Segment'] = df2['Segment'].astype(str) 
	df2['variable'] = df2['Segment']+'_'+df2['variable'] #samler view+variabel 
	del df2['Segment']
	df2['value'] = pd.to_numeric(df2.value, errors = 'coerce') #convert to numeric, og lav non-numeric strings til missing
	df3 = df2.groupby(['ID','variable'], as_index = False).agg({'value':'mean'})  # Midle dobbelt målinger!!        
	df3 = df3.pivot(index = 'ID', columns = 'variable', values = 'value' )
	#Færdig longitudinelt segmental
	df_segm_lax = df3
	#print(df_segm_lax.columns)
	########################################Load excel nr. 3 #####################################
	########################################Load excel nr. 3 #####################################
	df = pd.read_excel(filepath, sheet_name='SegmTime_LAX')
	cpr = df['ID']
	cpr = cpr.replace('#','',regex=True) 							#Panda har en replace( arguments .. ) funktion for dataframes. regex = True kræves for string
	cpr = cpr.replace('-','',regex=True)
	cpr = cpr.replace('migration','',regex=True)
	df['ID'] = cpr 
	del cpr 
	names = df.columns.str.strip(' ').str.replace(' ','_').str.replace('%','').str.replace('(','').str.replace(')','').str.replace('/','').str.replace('migration','').str.replace('migrationmigration','')
	df.columns = names
	#melt data
	df2 = pd.melt(df, id_vars=['ID','Segment'], value_name='value', var_name='variable')
	df2['variable'] = df2['variable'].astype(str)
	df2['Segment'] = df2['Segment'].astype(str) 
	df2['variable'] = 't_'+df2['Segment']+'_'+df2['variable'] #samler view+variabel 
	del df2['Segment']
	#print(df2)
	df2['value'] = pd.to_numeric(df2.value, errors = 'coerce') #convert to numeric, og lav non-numeric strings til missing
	df3 = df2.groupby(['ID','variable'], as_index = False).agg({'value':'mean'})  # Midle dobbelt målinger!!        
	df3 = df3.pivot(index = 'ID', columns = 'variable', values = 'value' )
	#Færdig longitudinelt segmental
	df_segm_time_lax = df3
	#print(df_segm_lax.columns)
	########################################Mergee data!!! #####################################
	#print(df_global_lax.index)

	df_global_lax.index.name = 'ID'   ## fix indices! 
	df_segm_time_lax.index.name = 'ID'
	df_segm_lax.index.name = 'ID'

	df_global_lax.reset_index(inplace=True)
	df_segm_lax.reset_index(inplace=True)
	df_segm_time_lax.reset_index(inplace=True)
	
	df_global_lax['ID'] = df_global_lax['ID'].astype(str)				#brug strings til CPR, fordi nogle gange er der bogstaver eller lign i ID 
	df_segm_time_lax['ID'] = df_segm_time_lax['ID'].astype(str)
	df_segm_lax['ID'] = df_segm_lax['ID'].astype(str) 


	df_combined = pd.merge(df_global_lax,df_segm_lax,on='ID')
	df_combined = pd.merge(df_combined,df_segm_time_lax,on='ID')

	return df_combined 



#Håndter fil indlæsning
def xml_find_files(path):
	#path = input('Enter path: ')		#Get path 
	os.chdir(path)						#change til path
	files = glob.glob('*.xml') 			#virker kun i working directory. har nu liste med files
	print('')
	print('')
	print('Number of files: {}'.format(len(files)))
	return files 




def parse_xml(path):
	mt = ET.parse(path)
	mr = mt.getroot()

	################# "Patient" node ekstrakt ####################
	list_vars = []
	list_values = []

	for i in mr.find('.//Patient'):
		list_vars.append(str(i.tag))
		list_values.append(str(i.text))

	df = pd.DataFrame(columns = list_vars)
	a_series = pd.Series(list_values, index = df.columns)
	df = df.append(a_series, ignore_index = True)


	# Since all xml files do not have all identifiers, only use the ones present in the dataframe
	possible_patient_identifier = ["FirstName", "PatientId", "LastName"]
	patient_identifiers = [patient_id for patient_id in possible_patient_identifier if patient_id in df.columns]

	df = df[patient_identifiers]
	df = df.rename(columns = {'PatientId':'cpr'})
	df_patient = df

	################# "Study" node ekstrakt ####################
	list_vars = []
	list_values = []

	for i in mr.find('.//Study'):
		list_vars.append(str(i.tag))
		list_values.append(str(i.text))

	df = pd.DataFrame(columns = list_vars)
	a_series = pd.Series(list_values, index = df.columns)
	df = df.append(a_series, ignore_index = True)
	drop_list = ['PregnancyOrigin','StudyInstanceUID','Series']
	df = df.drop(columns=drop_list)
	df['StudyDateTime'] = df['StudyDateTime'].astype(str)
	df['StudyDateTime'] = df['StudyDateTime'].str.slice(0,10)
	#print(df)
	df_study = df 

	################# "Parameter" node ekstrakt ####################
	list_vars = []
	list_values = []
	parameters = mr.findall('.//Parameter')

	if len(parameters) >= 1:

		#Lav start datasæt for parameter nr 1 
		for i in parameters[0]:
			list_vars.append(str(i.tag))
			list_values.append(str(i.text))

		df = pd.DataFrame(columns = list_vars)
		a_series = pd.Series(list_values, index = df.columns)
		df = df.append(a_series, ignore_index = True)
		#fjern første element fra liste med parameters
		parameters.pop(0) 

		for n in range(0,len(parameters)):
			#print('lol')
			list_vars = []
			list_values = []

			for i in parameters[n]:
				#print(i)
				list_vars.append(str(i.tag))
				list_values.append(str(i.text))

			df_temp = pd.DataFrame(columns = list_vars)
			a_series = pd.Series(list_values, index = df_temp.columns)
			df = df.append(a_series, ignore_index = True)

		df = df.rename(columns = {'DisplayUnit':'unit','ParameterName':'parameter','DisplayValue':'value'})
		df = df[['parameter','value']]
		df['value'] = abs(df.value.astype(float))
		df = df.dropna() #drop NA værdier 
		df = df.groupby(['parameter'], as_index = False).agg({'value':'mean'})
		df['fake_index'] = 1
		df = df.pivot(index = 'fake_index', values = 'value', columns = 'parameter' )
		df_parameter=df


		##### MERGE OG RETURN DATA ########
		df_patient['fake_index'] = 1 
		df_study['fake_index'] = 1 
		df_comb = pd.merge(df_patient,df_study,on='fake_index')
		df_comb = pd.merge(df_comb,df_parameter,on='fake_index')

		return df_comb


	if len(parameters) < 1:
		return 'Error'




######################################## GUI og main  #####################################

######################################### Knap kører excel proceess!! #####################################
def click(): 
	path = e1.get() #collect tekst i input felt 
	print(path)
	print('Processing XL files')
	files = excel_find_files(path)
	messagebox.showinfo('Files found','Number of XL files found: '+str(len(files))+'\n pres OK to merge')
	i=1 
	for p in files:
		print(i)
		print(p)
		if i==1:
			df = excel_process(p)
		if i!=1:
			df = df.append(excel_process(p))
		i += 1
	##Skriv til sti som filerne ligger i
	path_write=path+'/merged_XL.xlsx'
	global xl_path 
	xl_path = path_write
	print(path_write)
	df['cpr'] = df['ID']  	#### til at fixe irriterende index / not accesible column fejl 
	df.to_excel(path_write, engine = 'xlsxwriter')
	messagebox.showinfo('Result','Merge complete. Merged excel data available at path'+'\n'+path_write)
	global df_xl 
	df_xl = df
	#print(df)


######################################### Knap kører XML parse!! #####################################
def click2():
	path = e2.get() #collect tekst i input felt 
	print(path)
	files = xml_find_files(path)
	messagebox.showinfo('Files found','Number of XML files found: '+str(len(files))+'\n pres OK to merge')
	i=1
	for p in files:
		print(i)
		print(p)
		if i==1:
			df = parse_xml(p)
		if i!=1:
			j = parse_xml(p)
			if isinstance(j, pd.DataFrame):
				print(j)
				df = df.append(j)
		i += 1
	#Fix lidt småting
	names = df.columns.str.strip(' ').str.replace(' ','_').str.replace('%','').str.replace('(','').str.replace(')','').str.replace('/','').str.replace('migration','').str.replace('migrationmigration','')
	df.columns = names 
	cpr = df['cpr']
	cpr = cpr.replace('#','',regex=True) 							#Panda har en replace( arguments .. ) funktion for dataframes. regex = True kræves for string
	cpr = cpr.replace('-','',regex=True)
	cpr = cpr.replace('migration','',regex=True)
	df['cpr'] = cpr 
	del cpr 

	#tving til string 
	df['cpr'] = df['cpr'].astype(str) 

	##Skriv til sti som filerne ligger i
	path_write=path+'/Merged_XML.xlsx'
	global xml_path 
	xml_path = path_write
	print(path_write)
	df.to_excel(path_write, engine = 'xlsxwriter')
	messagebox.showinfo('Result','Merge complete. Merged XML data available at path'+'\n'+path_write)
	#print(df.columns)
	global df_xml 
	df_xml = df

######################################### Lav merged DTA #####################################
def click3():
	messagebox.showinfo('Create DTA','Create merged DTA file. The file will be placed in the same folder as the XL/XML files and is named "merged.dta". A merged XL file will also be created, named "merged.xlsx" Press OK to begin')
	#df_xml = pd.read_excel(xml_path, index_col = 0)
	#df_xl = pd.read_excel(xl_path, index_col = 0) #ikke nødvendig, datasæt XL og XML gemmes når den parser det andet i global variabel i funktionen
	
	global df_xl #tillad modifikation af globals i denne funktion 
	global df2_xl
	global df_xml 

	#hent prefix og rename alle columns hvis tilfældet 
	prefix = e3.get() 
	if isinstance(prefix, str):
		if prefix != '':
			lol = 'lol'
			print(lol)
			print(lol)
			print(lol)
			print(lol)
			df_xl = df_xl.add_prefix(prefix)
			temp = prefix+'cpr'
			df_xl = df_xl.rename(columns = {temp:'cpr'})
			#print(df_xl['cpr'])


	if df_xml is None:
		print('xml is none')
		#df_xl.to_stata('xls.dta', version=117)
		df_xl.to_excel('xls.xlsx', engine = 'xlsxwriter')
	elif df_xl is None: 
		print('xl is none')
		#df_xml.to_stata('XML.dta', version=117)
		df_xl.to_excel('xml.xlsx', engine = 'xlsxwriter')
	else:
		print('all data is here')
		#print(df_xl)
		df_xml['cpr'] = df_xml['cpr'].astype(str)
		df_xl['cpr'] = df_xl['cpr'].astype(str)
		#print(df_xl['cpr'])
		#print(df_xml['cpr'])
		df_merged = pd.merge(df_xl,df_xml, on = 'cpr', how='outer')
		print(df_merged)
		print(df_merged['cpr'])
		df_merged.index.name = 'index'
		print(df_merged.columns)
		df_merged.columns = df_merged.columns.str.strip()
		df_merged.to_excel('merged.xlsx', engine = 'xlsxwriter')
		#df_merged.to_stata('merged.dta', version=114, write_index = False, )

	for i in range(5):
		print('Final merge complete')



	

#main window##
window = tk.Tk() #lav master window 
window.title('Echopac XL data merger')
window.configure(background='black')

#Forklaringstekst ##
tk.Label(window, text='1: This program can merge Echopac speckle tracking XL file output and Echopac XML file output. It can ALSO create a final merged STATA DTA file (XML+XL speckle).', bg='black',fg='white').grid(row=0, sticky='W')
tk.Label(window, text='2: To use it, enter the file path for the folder containing the speckle tracking XL files or the Echopac XML files (same folder is OK) below and press "Merge"', bg='black',fg='white').grid(row=1, sticky='W')
tk.Label(window, text='3: The merged XL database "merged_XL.xlsx", merged XML database "Merged_XML.xlsx" and merged DTA "merged.dta" will be placed in the same folder as the files', bg='black',fg='white').grid(row=2, sticky='W')

#file path paste field ##
tk.Label(window, text="Paste file path to Echopac speckle tracking XL files in the box to the right").grid(row=4, sticky='E')
tk.Label(window, text="Paste file path to Echopac XML files in the box to the right").grid(row=6, sticky='E')
tk.Label(window, text='Add potential prefix to all variables except ID/CPR (ex. "LV_","RV_","LA_") Leave blank if none is desired').grid(row=8, sticky='E')
tk.Label(window, text="Press here to create merged DTA dataset (XML+XL merged in a STATA data file)").grid(row=9, sticky='E')

#path XL
e1 = tk.Entry(window)
e1.grid(row=4,column=1, sticky='W')

#path XML
e2 = tk.Entry(window)
e2.grid(row=6,column=1, sticky='W')

#prefix
e3 = tk.Entry(window)
e3.grid(row=8,column=1, sticky='W')

#Knap til submit
ttk.Button(window, text='Merge XL', width = 20, command=click).grid(row=5,column=0,sticky='E')
ttk.Button(window, text='Merge XML', width = 20, command=click2).grid(row=7,column=0,sticky='E')
#Lav stata datasæt
ttk.Button(window, text='Create merged DTA file', width = 20, command=click3).grid(row=9,column=1,sticky='W')

######################################## Kører programmet med GUI  #####################################
tk.mainloop()

#git test
#git test 2


	






