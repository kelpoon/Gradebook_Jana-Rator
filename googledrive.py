from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive

# Initialize GoogleDrive with the credentials
gauth = GoogleAuth()
gauth.LocalWebserverAuth()  # Authorize using a local web server
drive = GoogleDrive(gauth)

# ID of the folder you want to download files from
folder_id = '1tyeoX2ZgmgSW57cS8z-r4S_Z6hKPl2yo'

# Get a list of all files in the folder
file_list = drive.ListFile({'q': f"'{folder_id}' in parents and trashed=false"}).GetList()

# Download each file in the folder
for file in file_list:
    # Download the file to the current directory
    file.GetContentFile(file['title'])
    print(f"Downloaded: {file['title']}")

print("All files downloaded successfully!")
