# -*- coding: utf-8 -*-
"""
Rename multiple files within directory 

"""
# importing os module 
import os

# current path to directory with files
path = os.getcwd()+"/Rename_dir/"
  
# Function to rename multiple files 
def rename_files(path, new_name):
    ''' Renames all files in specified directory to one standard name
        with order number at the end, i.e new_name_1, new_name_2, etc.
        Takes:
        path - string - as path to directory with files
        new_name - string - as new file name for all files in directory'''
    # counter
    count = 1  
    for filename in os.listdir(path):
        # getting extention
        ext = "."+filename.split(".")[1]
        os.rename(path+filename, path+new_name+"_"+str(count)+ext)
        print("Renamed from {} to {}".format(filename, new_name+"_"+str(count)+ext))
        count += 1

rename_files(path, "New_name")
