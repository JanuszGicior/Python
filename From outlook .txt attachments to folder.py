import os
import win32com.client
from pathlib import Path

def download_attachments_and_move(inbox_folder_path, archive_folder_name):
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")

    inbox = namespace.GetDefaultFolder(6)  # 6 corresponds to the inbox folder

    # Retrieve the Archive folder or create it if it doesn't exist
    archive = None
    for folder in inbox.Folders:
        if folder.Name == archive_folder_name:
            archive = folder
            break

    if not archive:
        archive = inbox.Folders.Add(archive_folder_name)

    for item in inbox.Items:
        if item.Attachments.Count > 0:
            for attachment in item.Attachments:
                if attachment.FileName.endswith('.txt'):
                    desktop_path = str(Path.home() / "Desktop")
                    file_path = os.path.join(desktop_path, attachment.FileName)
                    attachment.SaveAsFile(file_path)

                    # Move the item to the Archive folder
                    item.Move(archive)

def main():
    inbox_folder_path = ""  # Add your inbox folder path if it's not the default
    archive_folder_name = "Archive"  # Change to your archive folder's name
    download_attachments_and_move(inbox_folder_path, archive_folder_name)

if __name__ == " 
