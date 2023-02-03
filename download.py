#provide the code on how to download files from sharepoint
from Office365_api import SharePoint
import re
import sys, os
from pathlib import PurePath 

# 1 args = SharePoint folder name. May include subfolders 
FOLDER_NAME = sys.argv[1]
#2 args = locate or remote folder_dest
FOLDER_DEST = sys.argv[2]
# 3 args = SharePoint file name. This is used when only one file is being downlaoded 
#if all files need to be downloaded, then set this value to 'None'
FILE_NAME = sys.argv[3]
# 4 args = SharePoiont file name pattern (filter attribute for folders containning files with different naming formats)
FILE_NAME_PATTERN = sys.argv [4]


def save_file(file_n, file_obj):
    file_dir_path = PurePath(FOLDER_DEST, file_n)
    with open(file_dir_path, 'wb') as f: #wb = write binary, can write different types of files
        f.write(file_obj)
        f.close()

def get_file(file_n, folder):
    file_obj = SharePoint().download_file(file_n, folder)
    save_file(file_n, file_obj)

#retrieve all the files in a specified folder
def get_files(folder):
    file_list = SharePoint()._get_files_list(folder)
    for file in file_list:
        get_file(file.name, folder)

def get_files_by_pattern(keyword, folder):
    file_list = SharePoint()._get_files_list(folder)
    for file in file_list:
        if re.search(keyword, file.name):
            get_file(file.name, folder)


if __name__ == '__main__':
    if FILE_NAME != 'None':
        get_file(FILE_NAME, FOLDER_NAME)
        print("Down")
    elif FILE_NAME_PATTERN != 'None':
        get_files_by_pattern(FILE_NAME_PATTERN, FOLDER_NAME)
        print("Load")
    else:
        get_files(FOLDER_NAME)

print("ALL_OKAY")