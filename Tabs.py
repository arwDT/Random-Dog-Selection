#import libraries
import pandas as pd
import numpy as np
from datetime import datetime
import PySimpleGUI as sg
import os
import sys
import re

#Drives and File Paths
Drive = 'C:\\'
Input = os.path.join(Drive, 'Python', 'RandomDogSelection', 'Input', 'Data-Extraction-Rosa-20200617.xlsx')

#Get column names for each sheet in data extract
def columns(sheet):
    cols = {}
    xl = pd.ExcelFile(sheet)
    for i in range(len(xl.sheet_names)):
        df = pd.read_excel(Input, sheet_name=i)
        cols[i] = df.columns
    return cols

#Return all surveys that cotain a dogid
def dog_surveys(cols):
    selection = []
    for key, value in cols.items():
        for j in value:
            if re.findall('Dog ID', str(j)):
                selection.append(key)
            elif re.findall('dogId', str(j)):
                selection.append(key)
    return selection

cols = (columns(Input))
surveys = (dog_surveys(cols))
print(surveys)









