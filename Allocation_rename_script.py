# -*- coding: utf-8 -*-
"""
Created on Sun Oct 20 12:21:57 2024

@author: Piyush
"""

# modules
import os
import pandas as pd

# get's input file path and folder path and returns the path
def get_paths():
    '''
    

    Returns
    -------
    excel_file_path : str
        path of specific excel file for this program.
    recording_folder_path : str
        path of the folder.

    '''
    excel_file_path = input("Enter the path of the Excel file to be used: ").strip().replace("\\", "/")
    if excel_file_path[0] == '"' and excel_file_path[-1] == '"':
            excel_file_path = excel_file_path[1:-1]
    recording_folder_path = input("Enter the path of the folder in which the recording are stored: ").strip().replace("\\", "/")
    if recording_folder_path[0] == '"' and recording_folder_path[-1] == '"':
            recording_folder_path = recording_folder_path[1:-1]
    return excel_file_path,recording_folder_path

def each_row_dict_list(excel_file_path:str):
    '''
    

    Parameters
    ----------
    excel_file_path : str
        path of specific excel file for this program.

    Returns
    -------
    rows_of_data_frame : list
        .

    '''
    counter = 1
    rows_of_data_frame = []
    while counter <=9:        
        df = pd.read_excel(excel_file_path)
        row_series = df.iloc[counter]
        row_dict = row_series.to_dict()
        rows_of_data_frame.append(row_dict)
        counter += 1            
    return rows_of_data_frame
     
# get's file names present in folder
def filenames_inside_folder(recording_folder_path:str)->list:
    '''
    

    Parameters
    ----------
    recording_folder_path : str
        DESCRIPTION.

    Returns
    -------
    recording_names : list
        list of file names present in folder.

    '''
    obj = os.scandir(recording_folder_path)   
    recording_names=[]
    for entry in obj :
        if entry.is_dir() or entry.is_file():
            recording_names.append(entry.name)
    return recording_names
  
# renaming the file names
def rename_script(recording_folder_path:str,rows_of_data_frame:list,recording_names:list):
    '''
    

    Parameters
    ----------
    recording_folder_path : str
        DESCRIPTION.
    rows_of_data_frame : list
        DESCRIPTION.
    recording_names : list
        DESCRIPTION.

    Returns
    -------
    None.

    '''
    for row in rows_of_data_frame:
        for value in row.values():
            for name in recording_names:
                if str(value) == name.split("_")[0]: # '==' symbol is not working so i have used in operator 
                    new_name = row["Name "]+".WAV"
                    os.rename(f"{recording_folder_path}\{name}",f"{recording_folder_path}\{new_name}")
            
            
if __name__=='__main__':
    excel_file_path,recording_folder_path = get_paths()
    rows_of_data_frame = each_row_dict_list(excel_file_path)
    recording_names = filenames_inside_folder(recording_folder_path)
    rename_script(recording_folder_path,rows_of_data_frame, recording_names)
    print("Renamed successfully")