# -*- coding: utf-8 -*-
"""
Created on Thu Jun 27 13:11:15 2019

@author: sagar_paithankar
"""
import os
import pandas as pd
#import xlrd

os.chdir(r'G:\Anaconda_CC\spyder\lightsource')
sheet_dict = pd.read_excel('Plant_Portfolio_till_July_2019.xls', sheet_name=None, usecols="A,F,H")

#plant_names = ['Hornacott', 'South Creake', 'Stragglethorpe Road', 'Binsted', \
#               'Charity Farm', 'Hill Harm', 'Safeguard Bradwall', 'Gelli Gron',\
#               'Bentley Estate 2', 'Cwrt Henllys', 'Morfa Farm', 'Middle Balbeggie', \
#               'Parsonage Wood', 'Fields Farm', 'Pen Rhiw', 'Meadow Farm (Thorpe Langton)', \
#               'Park Farm West', 'Sharland Farm', 'Millar Farm', 'Lough Road',\
#               'Moira Road', 'Knockcairn Road', 'Lough Road 2 (Hillside)', \
#               'Bolsovermoor Quarry', 'Belfast Road']

# =============================================================================
# ['PlantId', 'DA_MAE', 'ID_MAE', 'PlantId_other', 'R0_MAE',
#        'ID_MAE_other', 'PlantId_other', 'R0_MAE_other', 'ID_MAE_other',
#        'PlantId_other', 'R0_MAE_other', 'ID_MAE_other', 'PlantId_other',
#        'Capacity', 'R0_MAE_other', 'PlantId_other', 'R0_MAE_other',
#        'ID_MAE_other', 'PlantId_other', 'Capacity_other', 'R0_MAE_other',
#        'PlantId_other', 'Capacity_other', 'R0_MAE_other', 'PlantId_other',
#        'R0_MAE_other', 'ID_MAE_other', 'PlantId_other', 'R0_MAE_other',
#        'ID_MAE_other', 'PlantId_other', 'R0_MAE_other', 'ID_MAE_other',
#        'PlantId_other', 'Capacity_other', 'R0_MAE_other', 'PlantId_other',
#        'Capacity_other', 'R0_MAE_other', 'PlantId_other', 'Capacity_other',
#        'R0_MAE_other', 'PlantId_other', 'Capacity_other', 'R0_MAE_other',
#        'PlantId_other', 'Capacity_other', 'R0_MAE_other', 'PlantId_other',
#        'Capacity_other', 'R0_MAE_other', 'PlantId_other', 'Capacity_other',
#        'R0_MAE_other', 'PlantId_other', 'Capacity_other', 'R0_MAE_other',
#        'PlantId_other', 'Capacity_other', 'R0_MAE_other', 'PlantId_other',
#        'Capacity_other', 'R0_MAE_other', 'PlantId_other', 'Capacity_other',
#        'R0_MAE_other', 'PlantId_other', 'Capacity_other', 'R0_MAE_other',
#        'PlantId_other', 'Capacity_other', 'R0_MAE_other', 'PlantId_other',
#        'Capacity_other', 'R0_MAE_other']
# =============================================================================

full_table = pd.DataFrame()
grouping = pd.DataFrame()
for name, sheet in sheet_dict.items():
     a = pd.DataFrame(sheet)
     a['month'] = pd.to_datetime(a['DateTime'], format= "%Y-%m-%d").dt.month_name().str.slice(stop=3)
     a.drop(columns=['DateTime'], inplace=True)
     a.dropna(inplace=True)
     full_table = full_table.append(a)
     a = a.groupby(['month']).mean()
     grouping = grouping.append(a.T)

#plan = pd.Series(plant_names)    
#g = grouping.copy()
#g.drop(['PlantId'], inplace=True)
#g.reset_index(inplace=True)
#g.drop(columns = ['index'],inplace=True)
#p = pd.concat([plan, g], axis =1)
#p.set_index(0, inplace=True)
#import openpyxl
grouping.to_excel("Plant_Portfolio_Results.xlsx")


#p = full_table.describe()
#
#full_table.drop(columns=['PlantId', 'PlantId_other'], inplace=True)
#                         
#full_table.columns = plant_names
#
#full_table.to_excel("Plant_R0_MAE.xlsx")
#
#p.to_excel("Describe_Plant_R0_MAE.xlsx")