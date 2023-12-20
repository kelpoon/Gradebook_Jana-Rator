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


