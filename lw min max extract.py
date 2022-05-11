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
folder = '/Users/FA10041/'

#creates a list of strings of filenames for all .xls in a folder
for root, dirs, files in os.walk(folder):
    xlsfiles=[ _ for _ in files if _.endswith('.xls') ]

#function to extract optical data from powerview report
def read_excel_data(name):
    
    workbook = xlrd.open_workbook('C:' + folder + '/' + name)
    general_sheet = workbook.sheet_by_name('General')
    sheet1 = workbook.sheet_by_name('Sheet 1')
    sheet2 = workbook.sheet_by_name('Sheet 2')
    
    name = name
    
    general = general_sheet.cell(10, 1).value
       
    x_min = sheet1.cell(12, 2).value
    x_max = sheet1.cell(12, 3).value
    
    y_min = sheet2.cell(12, 2).value
    y_max = sheet2.cell(12, 3).value
    
    return [name, general, x_min, x_max, y_min, y_max]

#creates df of extracted values    
df = pd.concat([pd.DataFrame(read_pv_data(xlsfile)).T for xlsfile in xlsfiles])
df.reset_index(inplace=True)
df.drop('index', axis=1, inplace=True)

#adds column titles
df.columns = ['name', 'general', 'x_min', 'x_max', 'y_min', 'y_max']

#tidy general column
df['general'] = df['general'].str[:-1]
df['general'] = df['general'].astype(float)

#tidy name column and make type column
df['type'] = df['name'].str[:5]
df['name'] = df['name'].str[54:]
df['name'] = df['name'].str.rstrip('.xls')

#make sample no. and a z column
df['sample no.'] = df['name'].str[-3:]
df['sample no.'] = df['glass no.'].astype(int)
df['z'] = df['name'].str[-8:]
df['z'] = df['z'].str[:4]
df['z'] = df['z'].astype(float)

#adding true/false columns for info
def info_columns(info):
    df[info] = df['name'].str.contains(info, case=False)
    return df.head()

info = ['A', 'B']

for info in info:
    info_columns(info)
    
############################################################################################################
#swarmplots
############################################################################################################

#creating just one info column for the legend names
df_plot = df[['sample no.', 'type', 'z', 'x_min', 'x_max', 'y_min', 'y_max', 'A', 'B']].copy()

df_plot = df_plot.join((~df_plot[['A', 'B']]).add_prefix('inv_'))
df_plot[['A', 'B', 'inv_A', 'inv_B']]

info2 = ['A', 'B', 'inv_A', 'inv_B']

df_plot[info2] = df_plot[info2].where(df_plot != True, df_plot.columns.to_series(), axis=1)
df_plot[info2] = df_plot[info2].replace(False,'')

df_plot['info'] = df_plot[info2].apply(lambda row: ' '.join(row.values.astype(str)), axis=1)

df_plot = df_plot.drop(info2, 1)


#combining data for plots in one df so that all x and y values are in the same columns
df_min = df_plot[['type', 'z', 'info','x_min', 'y_min']]
df_max = df_plot[['type', 'info', 'x_max', 'y_max']]

df_plot_final = pd.concat((df_min, df_max), axis=0)

#can use this to filter by info
unique_infos = df_plot_final['info'].unique().tolist()


#plotting data

#swarm plot
sns.set_theme(style='whitegrid')
sns.set(rc={"figure.figsize":(14, 8)})
g = sns.swarmplot(x='info', y='x', data=df_plot_final.loc[df_plot_final['info'].isin(unique_info)], dodge=True, size=7)
loc, labels = plt.xticks()
g.set_xticklabels(labels, rotation=45, fontsize=14)
g.set_xlabel('Info', fontweight = 'bold', fontsize=16)
g.set_ylabel('x', fontweight = 'bold', fontsize=16)

#scatter of z vs x
sns.set(rc={"figure.figsize":(12, 8)})
g = sns.scatterplot(data=df_min, x='z', y='x', hue='Construction')
g.legend(fontsize='large', title_fontsize='40')
g.set_xlabel('z', fontweight = 'bold', fontsize=16)
g.set_ylabel('x', fontweight = 'bold', fontsize=16)

    


