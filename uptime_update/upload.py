from Office365_api import SharePoint
import environ
import re
import sys, os
from pathlib import PurePath

env = environ.Env() 
environ.Env.read_env()

# 1 args = Root Directory Path of files to upload
ROOT_DIR = r"C:\Users\Lesley Chingwena\Documents\python_scripts\Uptime\docs"
# 2 args = SharePoint folder name. May include subfolders to upload to
SHAREPOINT_FOLDER_NAME = env('sharepoint_folder')
# 3 args = File name pattern. Only upload files with this pattern
FILE_NAME_PATTERN = '2'

class Up:
    def upload_files(self, folder, keyword=None):
        file_list = self.get_list_of_files(folder)
        for file in file_list:
            if keyword is None or keyword == 'None' or re.search(keyword, file[0]):
                file_content = self.get_file_content(file[1])
                SharePoint().upload_file(file[0], SHAREPOINT_FOLDER_NAME, file_content)

    def get_list_of_files(self, folder):
        file_list = []
        folder_item_list = os.listdir(folder)
        for item in folder_item_list:
            item_full_path = PurePath(folder, item)
            if os.path.isfile(item_full_path):
                file_list.append([item, item_full_path])
        return file_list

    # read files and return the content of files
    def get_file_content(self, file_path):
        with open(file_path, 'rb') as f:
            return f.read()

    def upload(self, FILE_NAME_PATTERN):
        self.upload_files(ROOT_DIR, FILE_NAME_PATTERN)
        print("Upload")
