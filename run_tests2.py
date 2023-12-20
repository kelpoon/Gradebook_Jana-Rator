import unittest
from unittest.mock import patch, MagicMock
from pathlib import Path
import os
import os.path
import docx

# We want to test the functions defined in functions.py. Import them here

from functions import authenticate, escape_fname, search_folder, create_folder, calculateScoreFromHighlights

# Set an absolute path for root. This is based on the assumption that the
# source and test code are checked into the a folder path of the following
# form:
# <user-home-directory>/Desktop/<project-name>
#
# In case the source code is in another folder hierarchy, please adjust
# the below path

home=os.path.expanduser('~')
root = home+'/Desktop/Gradebook_Jana-Rator/drive_download'

# We will use one rubric file "Sample Student ..." to test the function
# process_rubric_files. We will mock the directory listing operations 
# and read the sample document. 

from functions import process_rubric_files
test_document='Sample Student - End Year Linear Algebra Grading Rubric .docx'


class TestFunctions(unittest.TestCase):

    def setUp(self):
        # Create the root directory if it doesnt exist 
        os.makedirs(root, exist_ok=True)

    # Add more tests for other functions as needed
    def test_calculateScoreFromHighlights_empty_list(self):
        self.assertEqual(calculateScoreFromHighlights([]), 0)

    def test_calculateScoreFromHighlights_single_highlight(self):
        self.assertEqual(calculateScoreFromHighlights([("text", 1)]), 1)

    def test_calculateScoreFromHighlights_multiple_highlights(self):
        highlights = [("text1", 0.5), ("text2", 1), ("text3", 0.75)]
        self.assertEqual(calculateScoreFromHighlights(highlights), 2.25)


#This test is for the main code that processes the rubric documents. The main code has to be reorganized into a function, process_rubric_files in order for the test to work. 
'''
class TestProcessRubricFiles(unittest.TestCase):

    @patch('functions.os.listdir')
    @patch('functions.Path.iterdir')
    #@patch('functions.docx.Document')
    @patch('functions.sheet.values().update')


    def test_process_rubric_files(self, mock_update, mock_document, mock_iterdir, mock_listdir):
        # Set up mocks
        mock_iterdir.return_value = [MagicMock()]
        mock_listdir.return_value = [test_document.docx]
        #mock_document.return_value = MagicMock()

        # Define the behavior of the mocked document
        # e.g., mock_document.return_value.tables = [...]

        # Call the function
        process_rubric_files()

        # Assertions
        # e.g., mock_update.assert_called_with(...)

'''

if __name__ == '__main__':
    unittest.main()


#Tests and code I made in November:
'''

import unittest
from unittest.mock import patch, MagicMock
import os

# Assuming the code is in a file named 'rubricReaderWriter.py'
from rubricReaderWriter import authenticate, escape_fname, search_folder, create_folder

# Set an absolute path for root based on your project structure
root = '/Users/amrrao/Desktop/Gradebook_Jana-Rator/drive_download'  # Adjust the value based on your actual setup

class TestRubricReaderWriter(unittest.TestCase):

    def setUp(self):
        # Create the necessary directory structure before each test
        os.makedirs(root, exist_ok=True)

    def test_authenticate(self):
        # Mock the GoogleAuth class to avoid actual authentication
        with patch('rubricReaderWriter.GoogleAuth') as mock_google_auth:
            gauth_instance = MagicMock()
            mock_google_auth.return_value = gauth_instance

            # Mock LoadCredentialsFile to return None (no saved credentials)
            gauth_instance.LoadCredentialsFile.return_value = None

            # Mock LocalWebserverAuth to avoid actual web server authentication
            gauth_instance.LocalWebserverAuth.return_value = None

            # Mock access_token_expired to simulate expired token
            gauth_instance.access_token_expired = True

            # Mock Refresh to avoid actual refreshing of credentials
            gauth_instance.Refresh.return_value = None

            # Mock Authorize to avoid actual authorization
            gauth_instance.Authorize.return_value = None

            # Mock SaveCredentialsFile to avoid actual saving of credentials
            gauth_instance.SaveCredentialsFile.return_value = None

            # Mock GoogleDrive class to avoid actual Google Drive connection
            with patch('rubricReaderWriter.GoogleDrive') as mock_google_drive:
                drive_instance = MagicMock()
                mock_google_drive.return_value = drive_instance

                # Call the function being tested
                result = authenticate()

                # Assert that the function returns the mocked GoogleDrive instance
                self.assertEqual(result, drive_instance)

    def test_escape_fname(self):
        # Test escape_fname function
        self.assertEqual(escape_fname("folder/name"), "folder_name")

    def test_search_folder(self):
        # Mock the authenticate function
        with patch('rubricReaderWriter.authenticate') as mock_authenticate:
            drive_instance = MagicMock()
            mock_authenticate.return_value = drive_instance

            # Mock the drive.ListFile function
            with patch.object(drive_instance, 'ListFile') as mock_list_file:
                # Mock the GetList function
                mock_get_list = MagicMock()
                mock_list_file.return_value.GetList.return_value = mock_get_list

                # Call the function being tested
                search_folder('folder_id', root)

                # Assert that the GetList function was called
                mock_list_file.assert_called_once_with({'q': "'%s' in parents and trashed=false" % 'folder_id'})
                mock_get_list.GetList.assert_called_once()

    def test_create_folder(self):
        # Test create_folder function
        path = root  # Use the root directory
        name = "test_folder"
        create_folder(path, name)
        self.assertTrue(os.path.exists(os.path.join(path, escape_fname(name))))

    # Add more tests for other functions as needed

if __name__ == '__main__':
    unittest.main()
Footer
'''
