import unittest
from book_generator.utils import clean_text_for_docx

class TestCleanTextForDocx(unittest.TestCase):
    def test_basic_replacement(self):
        text = "Hello\nWorld"
        expected = "Hello World"
        self.assertEqual(clean_text_for_docx(text), expected)

    def test_multiple_newlines(self):
        text = "Hello\n\nWorld"
        expected = "Hello World"
        self.assertEqual(clean_text_for_docx(text), expected)

    def test_surrounding_whitespace(self):
        text = "Hello \n   World"
        expected = "Hello World"
        self.assertEqual(clean_text_for_docx(text), expected)

    def test_no_newlines(self):
        text = "Hello World"
        expected = "Hello World"
        self.assertEqual(clean_text_for_docx(text), expected)

    def test_empty_string(self):
        text = ""
        expected = ""
        self.assertEqual(clean_text_for_docx(text), expected)

    def test_none_input(self):
        text = None
        expected = ""
        self.assertEqual(clean_text_for_docx(text), expected)

    def test_strips_outer_whitespace(self):
        text = "  Hello World  \n "
        expected = "Hello World"
        self.assertEqual(clean_text_for_docx(text), expected)

if __name__ == '__main__':
    unittest.main()
