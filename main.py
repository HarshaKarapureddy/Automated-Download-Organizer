from winreg import *
import os

#Different folders
organized_files = {"PDF folder":[".pdf"],
                   "PowerPoint folder":[".pptx", ".ppt"],
                   "Image folder":[".jpg", ".PNG", ".jpeg", ".png"],
                   "Application folder":[".exe", ".msi"],
                   "Zip folder":[".zip"],
                    "Word folder":[".doc", ".docx"],
                   "Excel folder":[".xlsx"],
                   "Video folder":[".mp4"],
                   "Audio folder":[".mp3"],
                   "Text folder":[".txt"],
                   "Miscellanious folder":[]}

with OpenKey(HKEY_CURRENT_USER, r'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders') as key:
    path_to_downloads = QueryValueEx(key, '{374DE290-123F-4565-9164-39C4925E467B}')[0]

path_to_downloads = os.path.join(path_to_downloads, "")
files = os.listdir(path_to_downloads)

#Creating folders
for folder in organized_files:
    folder_path = os.path.join(path_to_downloads, folder)
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)

#Organizing files
for file in files:
    file_path = os.path.join(path_to_downloads, file)
    if os.path.isfile(file_path):
        shifted = False
        file_extension = os.path.splitext(file)[1].lower()
        for folder, extensions in organized_files.items():
            if file_extension in extensions:
                try:
                    os.replace(file_path, os.path.join(path_to_downloads, folder, file))
                    shifted = True
                except Exception as e:
                    print(f"We had an error moving {file}: {e}")
                    break
        if not shifted:
            try:
                os.replace(file_path, os.path.join(path_to_downloads, "Miscellanious folder", file))
                shifted = True
            except Exception as e:
                print(f"We had an error moving {file}: {e}")
                break
    else:
        try:
            if file not in organized_files.keys():
                os.replace(file_path, os.path.join(path_to_downloads, "Miscellanious folder", file))
                shifted = True
        except Exception as e:
            print(f"We had an error moving {file}: {e}")
            break