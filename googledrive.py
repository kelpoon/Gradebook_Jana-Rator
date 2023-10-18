from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
import os

MIMETYPES = {
        # Drive Document files as MS dox
        'application/vnd.google-apps.document': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        # Drive Sheets files as MS Excel files.
        'application/vnd.google-apps.spreadsheet': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        # Drive presentation as MS pptx
        'application/vnd.google-apps.presentation': 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
        # see https://developers.google.com/drive/v3/web/mime-types
    }
EXTENSIONS = {
        'application/vnd.google-apps.document': '.docx',
        'application/vnd.google-apps.spreadsheet': '.xlsx',
        'application/vnd.google-apps.presentation': '.pptx'
}

def authenticate():
    gauth = GoogleAuth()

    # Tries to load saved client credentials
    gauth.LoadCredentialsFile("credentials.txt")
    
    if gauth.credentials is None:
        # Authenticate if they're not there
        gauth.LocalWebserverAuth()
    elif gauth.access_token_expired:
        # Refresh them if expired
        gauth.Refresh()
    else:
        # Initialize the saved credentials
        gauth.Authorize()

    # Save the current credentials to a file
    gauth.SaveCredentialsFile("credentials.txt")

    drive = GoogleDrive(gauth)
    return drive

def escape_fname(name):
    return name.replace('/','_')

def search_folder(folder_id, root):
    file_list = drive.ListFile({'q': "'%s' in parents and trashed=false" % folder_id}).GetList()
    for file in file_list:
        # print('title: %s, id: %s, kind: %s' % (file['title'], file['id'], file['mimeType']))
        # print(file)
        if file['mimeType'].split('.')[-1] == 'folder':
            foldername = escape_fname(file['title'])
            create_folder(root,foldername)
            search_folder(file['id'], '{}{}/'.format(root,foldername))
        else:
            download_mimetype = None
            filename = escape_fname(file['title'])
            filename = '{}{}'.format(root,filename)
            try:
                print('DOWNLOADING:', filename)
                if file['mimeType'] in MIMETYPES:
                    download_mimetype = MIMETYPES[file['mimeType']]

                    file.GetContentFile(filename+EXTENSIONS[file['mimeType']], mimetype=download_mimetype)
                else:
                    file.GetContentFile(filename)
            except:
                print('FAILED')
                f.write(filename+'\n')

def create_folder(path,name):
    os.mkdir('{}{}'.format(path,escape_fname(name)))

if __name__ == '__main__':
    drive = authenticate()

    f = open("failed.txt","w+")
    folder_id = '1tyeoX2ZgmgSW57cS8z-r4S_Z6hKPl2yo'
    root = 'drive_download'
    if not os.path.exists(root):
        os.makedirs(root)

    search_folder(folder_id,root+'/')
    f.close() 

