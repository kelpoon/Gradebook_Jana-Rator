# Functions #

# Drive related

def authenticate():
    """
    Authenticates the loads google drive for folder input
    :params: none
    :return: authorized Google Drive
    """
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
    """
    Simple function to reformat folder names
    :params: foldername
    :return: reformatted foldername
    """
    return name.replace('/','_')

def search_folder(folder_id, root):
    """
    Iterates through google drive folder until it finds google doc files to download
    If it finds relevant files, it downloads them to your local machine
    :param folder_id: URL ID of google drive folder
    :param root: computer root for download, i.e where the files should download
    :return: none, but downloads files
    """
    file_list = drive.ListFile({'q': "'%s' in parents and trashed=false" % folder_id}).GetList()
    for file in file_list:

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
    """
    Creates local file to house downloaded templates
    :param name: name of created folder
    :param path: computer root for download, i.e where the files should download
    :return: none
    """
    os.mkdir('{}{}'.format(path,escape_fname(name)))


# Terminal related

# for outputting pretty colors to the terminal
def convert_wd_color_index_to_termcolor(color_index):
    """
    Converts color ID to string colors if recognized, else returns white
    :param name: name of created folder
    :param path: computer root for download, i.e where the files should download
    :return: none
    """
    if (color_index == WD_COLOR_INDEX.BLUE):
        return "blue"
    if (color_index == WD_COLOR_INDEX.TURQUOISE):
        return "cyan"
    if (color_index == WD_COLOR_INDEX.GREEN):
        return "green"
    if (color_index == WD_COLOR_INDEX.BRIGHT_GREEN):
        return "light_green"
    if (color_index == WD_COLOR_INDEX.RED):
        return "red"
    if (color_index == WD_COLOR_INDEX.YELLOW):
        return "yellow"

    print("WARNING: Unrecognized color index: %s" % (color_index))
    return "white"


def calculateScoreFromHighlights(highlights):
    """
    Calculates total numerical score by summing highlights
    :param highlights: array of all highlights and corresponding scores
    :return: numerical score
    """
    score = 0
    for h in highlights:
        score += h[1]
    return score

# *************************************************************** #
