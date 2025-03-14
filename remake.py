import shutil
import os
from zipfile import ZipFile


def remake_files():
    files = []
    papka = os.getcwd()
    for file in os.listdir(papka):
        if file.endswith(".xlsx"):
            files.append(file)
    for file in files:
        tmp_folder = papka+'/tmp/convert_wrong_excel/'
        os.makedirs(tmp_folder, exist_ok=True)
        with ZipFile(file) as excel_container:
            excel_container.extractall(tmp_folder)
        wrong_file_path = os.path.join(tmp_folder, 'xl', 'SharedStrings.xml')
        correct_file_path = os.path.join(tmp_folder, 'xl', 'sharedStrings.xml')
        try:
            os.rename(wrong_file_path, correct_file_path) 
            os.remove(file)
            shutil.make_archive(file, 'zip', tmp_folder)
            os.rename(file+'.zip', file)
            print(file+' ready')
        except FileNotFoundError:
            pass
    return files

