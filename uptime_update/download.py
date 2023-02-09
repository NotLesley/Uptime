#This script download the required documents from sharepoint

from Office365_api import SharePoint
import environ
import re
import sys, os
from pathlib import PurePath 

env = environ.Env() 
environ.Env.read_env()

# SharePoint folder name where operating workbook is located. May include subfolders 
FOLDER_NAME = env('sharepoint_folder')
# location or remote folder_destination where workbook will be downloaded
FOLDER_DEST = r"C:\Users\Lesley Chingwena\Documents\python_scripts\Uptime\docs"
# 3 args = SharePoint file name. This is used when only one file is being downlaoded 
#if all files need to be downloaded, then set this value to 'None'
FILE_NAME = 'None'
# 4 args = SharePoiont file name pattern (filter attribute for folders containning files with different naming formats)
FILE_NAME_PATTERN = 'Uptime'

class Down:

    def save_file(self, file_n, file_obj):
        file_dir_path = PurePath(FOLDER_DEST, file_n)
        with open(file_dir_path, 'wb') as f: #wb = write binary, can write different types of files
            f.write(file_obj)
            f.close()

    def get_file(self, file_n, folder):
        file_obj = SharePoint().download_file(file_n, folder)
        self.save_file(file_n, file_obj)

    #retrieve all the files in a specified folder
    def get_files(self, folder):
        file_list = SharePoint()._get_files_list(folder)
        for file in file_list:
            self.get_file(file.name, folder)

    def get_files_by_pattern(self, keyword, folder):
        file_list = SharePoint()._get_files_list(folder)
        for file in file_list:
            if re.search(keyword, file.name):
                self.get_file(file.name, folder)

    def download(self, FILE_NAME): 
        if FILE_NAME != 'None':
            self.get_file(FILE_NAME, FOLDER_NAME)
            print("Download")
        elif FILE_NAME_PATTERN != 'None':
            self.get_files_by_pattern(FILE_NAME_PATTERN, FOLDER_NAME)
            print("Download")
        else:
            self.get_files(FOLDER_NAME)
