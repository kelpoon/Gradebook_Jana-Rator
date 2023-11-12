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
