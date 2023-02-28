from Office365_api import SharePoint
import re
import sys, os
from pathlib import PurePath

class Up:
    def upload_files(self, folder, dest, keyword=None):
        file_list = self.get_list_of_files(folder)
        for file in file_list:
            if keyword is None or keyword == 'None' or re.search(keyword, file[0]):
                file_content = self.get_file_content(file[1])
                SharePoint().upload_file(file[0], dest, file_content)

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

    def upload(self, file_name_pattern, root, destination):
        self.upload_files(root, destination, file_name_pattern)
