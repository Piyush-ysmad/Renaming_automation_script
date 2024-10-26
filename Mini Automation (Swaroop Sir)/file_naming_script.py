# -*- coding: utf-8 -*-
"""
Created on Fri Oct 11 21:29:39 2024

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
        get's path as input of the specific excel file for this program.
    recording_folder_path : str
        get's path as input of the folder.

    '''
    excel_file_path = input("Enter the path of the Excel file to be used: ").strip().replace("\\", "/")
    if excel_file_path[0] == '"' and excel_file_path[-1] == '"':
            excel_file_path = excel_file_path[1:-1]
    recording_folder_path = input("Enter the path of the folder in which the recording are stored: ").strip().replace("\\", "/")
    if recording_folder_path[0] == '"' and recording_folder_path[-1] == '"':
            recording_folder_path = recording_folder_path[1:-1]
    return excel_file_path,recording_folder_path

# get's all name from Name column
# get's all numbers from phone no. column
def column_selector(excel_file_path:str)->list:
    '''
    

    Parameters
    ----------
    excel_file_path : str
        path of excel file.

    Returns
    -------
    name_list : list
        All names in list format.
    phone_list : list
        All numbers in list format.

    '''
    df = pd.read_excel(excel_file_path) 
    name_list = df['Name'].tolist()
    
    df = pd.read_excel(excel_file_path)
    phone_list = df['Phone No'].tolist()
    
    return name_list,phone_list      

# create dictionary with keys as phone no.s and values as names
def dictionay_creator(name_list:list,phone_list:list)->dict:
    '''
    

    Parameters
    ----------
    name_list : list
        DESCRIPTION.
    phone_list : list
        DESCRIPTION.

    Returns
    -------
    name_and_phone_dictionary : dict
        keys as phone no.s and values as names.

    '''
    total_name = len(name_list)
    name_and_phone_dictionary = {}    
    for i in range(0,total_name):
        name_and_phone_dictionary[phone_list[i]] = name_list[i]
    return name_and_phone_dictionary

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
def rename_script(recording_folder_path:str,name_and_phone_dictionary:dict,recording_names:list):
    '''
    

    Parameters
    ----------
    recording_folder_path : str
        DESCRIPTION.
    name_and_phone_dictionary : dict
        DESCRIPTION.
    recording_names : list
        DESCRIPTION.

    Returns
    -------
    None.

    '''
    for i in name_and_phone_dictionary.keys():
        for j in recording_names:
            if str(i) in j.split("_")[0]: # '==' symbol is not working so i have used in operator 
                new_name = name_and_phone_dictionary[i]+".WAV"
                os.rename(f"{recording_folder_path}\{j}",f"{recording_folder_path}\{new_name}")
            
            
if __name__=='__main__':
    excel_file_path,recording_folder_path = get_paths()
    name_list,phone_list = column_selector(excel_file_path)
    name_and_phone_dictionary = dictionay_creator(name_list,phone_list)
    recording_names = filenames_inside_folder(recording_folder_path)
    rename_script(recording_folder_path,name_and_phone_dictionary, recording_names)