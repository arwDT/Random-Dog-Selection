'''
Future impovements:
Iterate through tabs, starting at most recect tab and working backwards so as not to include dogs removed from study
'''



#import libraries
import pandas as pd
import numpy as np
from datetime import datetime
import PySimpleGUI as sg
import os
import sys

def Repeat(x):
    _size = len(x)   #Get size of list
    repeated = []    #Initiate list for duplicate values
    for i in range(_size):  #Iterate over array the size of the user id
        k = i + 1  #range + 1
        for j in range(k, _size):
            if x[i] == x[j] and x[i] not in repeated:
                repeated.append(x[i])
    return repeated

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

#Drives and File Paths
#Drive = 'C:\\'
#Output = os.path.join(Drive, 'Python', 'RandomDogSelection', 'Output')

#Get current datetime
now = datetime.now()
dt_string = now.strftime("%Y%m%d%H%S")

#Import extract
df = pd.read_excel(file_path)

#Create copy of extract
df_1 = df

#Get List of userid's
user_ids = df['User ID'].tolist()

#Get repeated id's in user_ids
dupes = Repeat(user_ids)

#Remove all dupes from orginal df
for dupe in dupes:
    df = df[df['User ID'] != dupe]

#For each duplicate userid in df_1 (copy), randomly select one dog and append to df containing non-duplicate ids in
#filtered df
for dupe in dupes:
    filt_1 = df_1[df_1['User ID'] == dupe]
    length = len(filt_1)
    rand = np.random.randint(0, length)
    filt_2 = filt_1.iloc[[rand]]
    df = pd.concat([df, filt_2])

#Input file name - maybe change to a GUI in futute?
#file_name = input('Please Name Your File: ')
#Output to excel
#df.to_excel(output_path + '\{}_{}.xlsx'.format(file_name, dt_string), index=False)

print("Random Selection Complete")















