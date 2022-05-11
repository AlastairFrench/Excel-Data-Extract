# -*- coding: utf-8 -*-
"""
Created on Tue Apr  5 09:35:25 2022

@author: FA10041
"""

#import required libraries
import pandas as pd
import numpy as np
import xlrd
import os
import seaborn as sns
import matplotlib.pyplot as plt

#path to folder in which the files you're extracting are kept - must be /
folder = '/Users/FA10041/Documents/Camera Windows/Powerview/Lightweight/reports/'

#creates a list of strings of filenames for all .xls in a folder
for root, dirs, files in os.walk(folder):
    xlsfiles=[ _ for _ in files if _.endswith('.xls') ]

#function to extract optical data from powerview report
def read_pv_data(name):
    
    workbook = xlrd.open_workbook('C:' + folder + '/' + name)
    general_sheet = workbook.sheet_by_name('General')
    h_sheet = workbook.sheet_by_name('Camera Viewports 2')
    v_sheet = workbook.sheet_by_name('Camera Viewports 1')
    
    name = name
    
    angle = general_sheet.cell(10, 1).value
       
    h_min = h_sheet.cell(12, 2).value
    h_max = h_sheet.cell(12, 3).value
    
    v_min = v_sheet.cell(12, 2).value
    v_max = v_sheet.cell(12, 3).value
    
    return [name, angle, h_min, h_max, v_min, v_max]

#creates df of extracted values    
df = pd.concat([pd.DataFrame(read_pv_data(xlsfile)).T for xlsfile in xlsfiles])
df.reset_index(inplace=True)
df.drop('index', axis=1, inplace=True)

#adds column titles
df.columns = ['name', 'angle', 'h_min', 'h_max', 'v_min', 'v_max']

#tidy angle column
df['angle'] = df['angle'].str[:-1]
df['angle'] = df['angle'].astype(float)

#tidy name column and make model column
df['model'] = df['name'].str[:5]
df['name'] = df['name'].str[54:]
df['name'] = df['name'].str.rstrip('.xls')

#make glass no. and xc column
df['glass no.'] = df['name'].str[-3:]
df['glass no.'] = df['glass no.'].astype(int)
df['xc'] = df['name'].str[-8:]
df['xc'] = df['xc'].str[:4]
df['xc'] = df['xc'].astype(float)

#adding true/false columns for construction
def construct_columns(construct):
    df[construct] = df['name'].str.contains(construct, case=False)
    return df.head()

constructs = ['S4', '18']

for construct in constructs:
    construct_columns(construct)
    
############################################################################################################
#swarmplots
############################################################################################################

#creating just one construction column for the legend names
df_plot = df[['glass no.', 'model', 'xc', 'h_min', 'h_max', 'v_min', 'v_max', 'S4', '18']].copy()

df_plot = df_plot.join((~df_plot[['S4', '18']]).add_prefix('inv_'))
df_plot[['S4', '18', 'inv_S4', 'inv_18']]
df_plot.rename(columns={'S4': 'S2+S4 IRR', '18': '1.8/1.4mm', 'inv_S4': 'S2 only', 'inv_18': '2.1/1.6mm'}, inplace=True)

constructs2 = ['S2+S4 IRR', 'S2 only', '1.8/1.4mm', '2.1/1.6mm']

df_plot[constructs2] = df_plot[constructs2].where(df_plot != True, df_plot.columns.to_series(), axis=1)
df_plot[constructs2] = df_plot[constructs2].replace(False,'')

df_plot['construction'] = df_plot[constructs2].apply(lambda row: ' '.join(row.values.astype(str)), axis=1)

df_plot = df_plot.drop(constructs2, 1)


#combining data for plots in one df
df_min = df_plot[['model', 'xc', 'construction','h_min', 'v_min']]
df_min.columns = ['Model', 'Cross Curve (mm)', 'Construction', 'Horizontal Optical Power (mDpt)', 'Vertical Optical Power (mDpt)']
df_max = df_plot[['model', 'construction', 'h_max', 'v_max']]
df_max.columns = ['Model', 'Construction', 'Horizontal Optical Power (mDpt)', 'Vertical Optical Power (mDpt)']

df_plot_final = pd.concat((df_min, df_max), axis=0)

#can use this to filter by consturction
unique_constructs = df_plot_final['Construction'].unique().tolist()


#plotting data
sns.set_theme(style='whitegrid')
sns.set(rc={"figure.figsize":(14, 8)})
g = sns.swarmplot(x='Construction', y='Vertical Optical Power (mDpt)', data=df_plot_final.loc[df_plot_final['Construction'].isin(unique_constructs)], dodge=True, size=7)
loc, labels = plt.xticks()
g.set_xticklabels(labels, rotation=45, fontsize=14)
g.set_xlabel('Construction', fontweight = 'bold', fontsize=16)
g.set_ylabel('Vertical Optical Power (mDpt)', fontweight = 'bold', fontsize=16)

sns.set_theme(style='whitegrid')
sns.set(rc={"figure.figsize":(14, 8)})
g = sns.swarmplot(x='Construction', y='Horizontal Optical Power (mDpt)', data=df_plot_final.loc[df_plot_final['Construction'].isin(unique_constructs)], dodge=True, size=7)
loc, labels = plt.xticks()
g.set_xticklabels(labels, rotation=45, fontsize=14)
g.set_xlabel('Construction', fontweight = 'bold', fontsize=16)
g.set_ylabel('Horizontal Optical Power (mDpt)', fontweight = 'bold', fontsize=16)

############################################################################################################
#scatter of xc vs min vert
############################################################################################################
sns.set(rc={"figure.figsize":(12, 8)})
g = sns.scatterplot(data=df_min, x='Cross Curve (mm)', y='Vertical Optical Power (mDpt)', hue='Construction')
g.legend(fontsize='large', title_fontsize='40')
g.set_xlabel('Cross Curve (mm)', fontweight = 'bold', fontsize=16)
g.set_ylabel('Minimum Vertical Optical Power (mDpt)', fontweight = 'bold', fontsize=16)

############################################################################################################
#vert data analysis
############################################################################################################
average_vert_min = {}
for i in range(len(unique_constructs)):
    average_vert_min[i] = {unique_constructs[i]:df_min['Vertical Optical Power (mDpt)'].loc[df_min['Construction'] == unique_constructs[i]].mean()}

std_vert_min = {}
for i in range(len(unique_constructs)):
    std_vert_min[i] = {unique_constructs[i]:df_min['Vertical Optical Power (mDpt)'].loc[df_min['Construction'] == unique_constructs[i]].std()}
    


