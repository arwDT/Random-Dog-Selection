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
file_path = os.path.join(Drive, 'Python', 'RandomDogSelection', 'Input', 'Data-Extraction-Rosa-20200617.xlsx')
output_path = os.path.join(Drive, 'Python', 'RandomDogSelection', 'Output')
file_name = 'Test_R_to_L'

#Get column names for each sheet in data extract
def columns(sheet):
    cols = {}
    xl = pd.ExcelFile(sheet)
    for i in range(len(xl.sheet_names)):
        df = pd.read_excel(sheet, sheet_name=i)
        cols[i] = df.columns
    return cols

#Return all surveys that cotain a dogid
def dog_surveys(cols):
    selection = []
    for key, value in cols.items():
        if len(re.findall('dogId', str(value))) > 0:
            selection.append(key)
    return selection

def non_dog_surveys(cols):
    selection = []
    for key, value in cols.items():
        if len(re.findall('dogId', str(value))) == 0:
            selection.append(key)
    return selection

#Function to find all repeated user_ids
def Repeat(x):
    _size = len(x)   #Get size of list
    repeated = []    #Initiate list for duplicate values
    for i in range(_size):  #Iterate over array the size of the user id
        k = i + 1  #range + 1
        for j in range(k, _size):
            if x[i] == x[j] and x[i] not in repeated:
                repeated.append(x[i])
    return repeated

#Function to return one dog per user_id
def random_select(df):
    # Create copy of extract
    df_1 = df

    # Get List of userid's
    user_ids = df['userId'].tolist()

    # Get repeated id's in user_ids
    dupes = Repeat(user_ids)

    # Remove all dupes from orginal df
    for dupe in dupes:
        df = df[df['userId'] != dupe]

    # For each duplicate userid in df_1 (copy), randomly select one dog and append to df containing non-duplicate ids in
    # filtered df
    for dupe in dupes:
        filt_1 = df_1[df_1['userId'] == dupe]
        length = len(filt_1)
        rand = np.random.randint(0, length)
        filt_2 = filt_1.iloc[[rand]]
        df = pd.concat([df, filt_2])

    return df

#Get current datetime
now = datetime.now()
dt_string = now.strftime("%Y%m%d%H%S")

#Get tabs from survey containing dogids
dog_surveys = (dog_surveys(columns(file_path)))
non_dog_surveys = (non_dog_surveys(columns(file_path)))

#Reverse list of dog surveys
dog_surveys_rl = dog_surveys[::-1]

#Get tab names from original file
xl = pd.ExcelFile(file_path)
names = xl.sheet_names

print(dog_surveys_rl)


#Create writer
writer = pd.ExcelWriter(output_path + '\{}_{}.xlsx'.format(file_name, dt_string), engine='xlsxwriter')
writer_f_1_rnd = pd.ExcelWriter(output_path + '\{}_{}.xlsx'.format('Pt 1', dt_string), engine='xlsxwriter')
writer_f_2 = pd.ExcelWriter(output_path + '\{}_{}.xlsx'.format('Pt 2', dt_string), engine='xlsxwriter')


#Create empty list of dog
dog_id_f = []
user_id_f = []
for survey in dog_surveys_rl:
    #Get the most recent survey
    sheet = pd.read_excel(file_path, sheet_name=survey)
    #Get all id's not previously selected
    sheet_fd_1 = sheet[-sheet['userId'].isin(user_id_f)]
    # Randomly select dogids for multiple owner ids from remaining userid's
    Sheet_f_1_rnd = random_select(sheet_fd_1)
    #Get previously selected ids
    sheet_f_2 = sheet[sheet['dogId'].isin(dog_id_f)]
    #Combine previously selected and newly selected ids
    sheet_f = pd.concat([Sheet_f_1_rnd, sheet_f_2])
    # Get list of user_ids associated with filtered dog ids
    for i in Sheet_f_1_rnd['userId'].tolist():
        user_id_f.append(i)
    # Add these id's to dog_id_f
    for j in Sheet_f_1_rnd['dogId'].tolist():
        dog_id_f.append(j)
    #Write to spreadsheet
    sheet_f.to_excel(writer, sheet_name=names[survey], index=False)
    Sheet_f_1_rnd.to_excel(writer_f_1_rnd, sheet_name=names[survey], index=False)
    sheet_f_2.to_excel(writer_f_2, sheet_name=names[survey], index=False)


writer.save()
writer_f_1_rnd.save()
writer_f_2.save()
print("Random Selection Complete")