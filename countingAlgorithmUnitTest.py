import unittest
from unittest.mock import patch, MagicMock
from your_script import escape_fname, convert_wd_color_index_to_termcolor, calculateScoreFromHighlights, count_highlights_in_range

class TestFunctions(unittest.TestCase):

    def test_escape_fname(self):
        input_name = 'my/folder/name'
        expected_output = 'my_folder_name'
        self.assertEqual(escape_fname(input_name), expected_output)

        input_name = 'foldername'
        self.assertEqual(escape_fname(input_name), input_name)

    def test_convert_wd_color_index_to_termcolor(self):
        color_index = 4  # Sample color index
        expected_output = "blue"
        self.assertEqual(convert_wd_color_index_to_termcolor(color_index), expected_output)

        # Add more test cases for other color indices

    def test_calculateScoreFromHighlights(self):
        highlights = [("text1", 2), ("text2", 3), ("text3", 1)]
        expected_score = 6
        self.assertEqual(calculateScoreFromHighlights(highlights), expected_score)

        # Test with an empty highlights list
        empty_highlights = []
        self.assertEqual(calculateScoreFromHighlights(empty_highlights), 0)

    # It's challenging to test count_highlights_in_range without having a Document object or mocking docx functionality
    # This test demonstrates how you might structure a test using mock, but it won't work directly without adjustments.
    @patch('your_script.docx.Document')
    def test_count_highlights_in_range(self, mock_document):
        fake_document = MagicMock()
        fake_table = MagicMock()
        fake_table.rows = [
            MagicMock(cells=[MagicMock(paragraphs=[MagicMock(text="text1")]),
                              MagicMock(paragraphs=[MagicMock(text="text2")])]),
            # Add more rows and cells as needed for your test cases
        ]
        fake_document.tables = [fake_table]

        mock_document.return_value = fake_document

        # Now call count_highlights_in_range with the fake document and test its behavior
        # count_highlights_in_range(document, table_num=0, start_row=1, end_row=1, start_column=1, end_column=1, darkColor=4, lightColor=11)
        result = count_highlights_in_range(fake_document, table_num=0, start_row=1, end_row=1, start_column=1, end_column=1, darkColor=4, lightColor=11)

        # Add assertions based on the expected behavior of count_highlights_in_range

        # Example assertion:
        self.assertEqual(result, (0, 0, 0))  # Replace with your expected output

if __name__ == '__main__':
    unittest.main()
