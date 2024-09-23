import mimetypes
import os
import time
from winreg import *

from watchdog.events import FileSystemEvent, FileSystemEventHandler
from watchdog.observers import Observer


class Download_Orgainzer(FileSystemEventHandler):

    def __init__(self):
        self.organized_files = {
            "PDF folder": ["application/pdf"],
            "PowerPoint folder": [
                "application/vnd.ms-powerpoint",
                "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            ],
            "Image folder": [
                "image/jpeg",
                "image/png",
                "image/gif",
                "image/bmp",
                "image/webp",
                "image/svg+xml",
                "image/tiff",
                "image/x-icon",
            ],
            "Application folder": ["application/x-msdownload"],
            "Zip folder": [
                "application/zip",
                "application/x-zip-compressed",
                "application/x-rar-compressed",
                "application/x-7z-compressed",
                "application/x-tar",
            ],
            "Word folder": [
                "application/msword",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            ],
            "Excel folder": [
                "application/vnd.ms-excel",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            ],
            "Video folder": [
                "video/mp4",
                "video/mpeg",
                "video/webm",
                "video/ogg",
                "video/x-msvideo",
                "video/quicktime",
                "video/x-flv",
            ],
            "Audio folder": [
                "audio/mpeg",
                "audio/wav",
                "audio/ogg",
                "audio/aac",
                "audio/webm",
                "audio/flac",
            ],
            "Text folder": ["text/plain"],
            "Miscellaneous folder": [],
        }
        # adding unidentified types for classification
        mimetypes.add_type("application/x-msdownload", ".msi")
        mimetypes.add_type("application/vnd.ms-excel", ".xlsm")
        mimetypes.add_type("application/x-7z-compressed", ".7z")

        # initializing download path and folders
        self.path_to_download = self.get_download_path()
        self.folder_creator()

        # initializing Observer for the event handler
        self.observer = Observer()

        self.moved_files = set()

    def get_download_path(self):
        # getting user's download path
        with OpenKey(
            HKEY_CURRENT_USER,
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders",
        ) as key:
            return QueryValueEx(key, "{374DE290-123F-4565-9164-39C4925E467B}")[0]

    def folder_creator(self):
        for folder in self.organized_files:
            folder_path = os.path.join(self.path_to_download, folder)
            if not os.path.exists(folder_path):
                os.makedirs(folder_path)

    def on_modified(self, event):
        time.sleep(10)
        self.organize_files()

    def on_created(self, event):
        time.sleep(10)
        self.organize_files()

    def organize_files(self):
        for file in os.listdir(self.path_to_download):
            file_path = os.path.join(self.path_to_download, file)
            if os.path.isfile(file_path):
                shifted = False

                mime_type, _ = mimetypes.guess_type(file_path)

                if mime_type is not None:
                    for folder, mime_types in self.organized_files.items():
                        if mime_type in mime_types:
                            self.move_file(file_path, folder, file)
                            shifted = True
                            break

                if not shifted:
                    self.move_file(file_path, "Miscellaneous folder", file)

            elif os.path.isdir(file_path):
                # Ignore known folders (already created by the script)
                if file not in self.organized_files.keys():
                    self.move_file(file_path, "Miscellaneous folder", file)

    def move_file(self, file_path, folder, file):
        if file in self.moved_files:
            print("File has already been moved. Skipping")
            return

        destination = os.path.join(self.path_to_download, folder, file)

        if not os.path.exists(file_path):
            print(f"File {file} no longer exists.")
            return

        try:
            os.replace(file_path, destination)
            print(f"Moved {file} to {folder}")
            # Add the file to the set of moved files
            self.moved_files.add(file)
        except Exception as e:
            print(
                f"We had an error moving {file}: File might not exist or is downloading (wait if so)."
            )

    def app(self):
        self.observer.schedule(self, self.path_to_download, recursive=False)
        self.observer.start()
        print("organization started")

        try:
            while True:
                time.sleep(60)
                self.organize_files()

        except KeyboardInterrupt:
            print("Stopping the observer and terminating the program")
            self.observer.stop()
        # finishes any task that was still in the works
        finally:
            self.observer.join()
            print("Program terminated")


if __name__ == "__main__":
    organizer = Download_Orgainzer()
    organizer.app()
