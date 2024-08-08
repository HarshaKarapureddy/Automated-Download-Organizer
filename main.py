from winreg import *
import os

organized_files = ["PDF folder", "PowerPoint folder", "Image folder", "Application folder", "Zip folder",
                    "Word folder", "Excel folder", "Video folder", "Audio folder", "Text folder",
                   "Miscellanious folder"]

with OpenKey(HKEY_CURRENT_USER, r'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders') as key:
    path_to_downloads = QueryValueEx(key, '{374DE290-123F-4565-9164-39C4925E467B}')[0]

path_to_downloads = path_to_downloads + '\\'


files = os.listdir(path_to_downloads)


for target in organized_files:
    if target not in files:
        os.makedirs(os.path.join(path_to_downloads, target))

for file in files:
    if file not in organized_files:
        if(file.endswith(".pdf")):
            os.replace(path_to_downloads + file , path_to_downloads + "PDF folder\\" + file)
        elif (file.endswith(".pptx") or file.endswith(".ppt")):
            os.replace(path_to_downloads + file, path_to_downloads + "PowerPoint folder\\" + file)
        elif (file.endswith(".jpg") or file.endswith(".PNG") or file.endswith(".jpeg") or file.endswith(".png")):
            os.replace(path_to_downloads + file, path_to_downloads + "Image folder\\" + file)
        elif (file.endswith(".exe") or file.endswith(".msi")):
            os.replace(path_to_downloads + file, path_to_downloads + "Application folder\\" + file)
        elif (file.endswith(".zip")):
            os.replace(path_to_downloads + file, path_to_downloads + "Zip folder\\" + file)
        elif (file.endswith(".doc") or file.endswith(".docx")):
            os.replace(path_to_downloads + file, path_to_downloads + "Word folder\\" + file)
        elif (file.endswith(".xlsx")):
            os.replace(path_to_downloads + file, path_to_downloads + "Excel folder\\" + file)
        elif (file.endswith(".mp4")):
            os.replace(path_to_downloads + file, path_to_downloads + "Video folder\\" + file)
        elif (file.endswith(".mp3")):
            os.replace(path_to_downloads + file, path_to_downloads + "Audio folder\\" + file)
        elif (file.endswith(".txt")):
            os.replace(path_to_downloads + file, path_to_downloads + "Text folder\\" + file)
        else:
            os.replace(path_to_downloads + file, path_to_downloads + "Miscellanious folder\\" + file)