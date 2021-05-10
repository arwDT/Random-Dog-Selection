#import libraries
import pandas as pd
import numpy as np
from datetime import datetime
import PySimpleGUI as sg
import os
import sys
import re

#Drives and File Paths
#Drive = 'C:\\'
#file_path = os.path.join(Drive, 'Python', 'RandomDogSelection', 'Input', 'Data-Extraction-Rosa-20200617.xlsx')


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

#Create GUI to input file path
sg.theme('DarkAmber')
# All the stuff inside window.
layout = [[sg.Text('Select File:')],
          [sg.In(),sg.FileBrowse()],
          [sg.Text('Select Output Folder:')],
          [sg.In(),sg.FolderBrowse()],
          [sg.Text('Name File:')],
          [sg.In()],
          [sg.Button('Ok'), sg.Button('Cancel')]]

# Create the Window
window = sg.Window('Window Title', layout)
# Event Loop to process "events" and get the "values" of the inputs
while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Cancel': # if user closes window or clicks cancel
        window.close()
        sys.exit()
        break
    if event == 'Ok': #if user clicks okay
        file_path = values.get(0)
        output_path = values.get(1)
        file_name = values.get(2)
        window.close()
        break

#Get current datetime
now = datetime.now()
dt_string = now.strftime("%Y%m%d%H%S")

#Get tabs from survey containing dogids
surveys = (dog_surveys(columns(file_path)))
non_dog_surveys = (non_dog_surveys(columns(file_path)))

#Get first tab with dogids
Sheet1 = pd.read_excel(file_path, sheet_name=surveys[0])

#Randomly select dogids for multiple owner ids
Sheet1_f = random_select(Sheet1)

#Get list of dogids from filtered extract
dog_id_f = Sheet1_f['dogId'].tolist()

#Get tab names from original file
xl = pd.ExcelFile(file_path)
names = xl.sheet_names

#Create writer
writer = pd.ExcelWriter(output_path + '\{}_{}.xlsx'.format(file_name, dt_string), engine='xlsxwriter')

for non_dog_survey in non_dog_surveys:
    sheet = pd.read_excel(file_path, sheet_name=non_dog_survey)
    sheet.to_excel(writer, sheet_name=names[non_dog_survey], index=False)

#filter each tab and add it to a new output dataframe
for survey in surveys:
    sheet = pd.read_excel(file_path, sheet_name=survey)
    sheet_f = sheet[sheet['dogId'].isin(dog_id_f)]
    sheet_f.to_excel(writer, sheet_name=names[survey], index=False)

writer.save()
print("Random Selection Complete")












