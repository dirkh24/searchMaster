import unittest
from unittest.mock import patch
from io import StringIO
from main import search_files, format_results
import os
import sys


class TestSearchProgram(unittest.TestCase):
    def setUp(self):
        # Erstelle temporäres Verzeichnis und Dateien für Tests
        self.test_dir = 'test_dir'
        is_exist = os.path.exists(self.test_dir)
        if not is_exist:
            os.makedirs(self.test_dir)

        with open(os.path.join(self.test_dir, 'test_file.txt'), 'w') as f:
            f.write("Test content with keyword")

        # Erstelle eine temporäre Excel-Datei für Tests
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws['A1'] = "Excel keyword test"
        wb.save(os.path.join(self.test_dir, 'test_file.xlsx'))

    def tearDown(self):
        # Lösche temporäres Verzeichnis nach Tests
        import shutil
        shutil.rmtree(self.test_dir)

    def test_search_files_txt(self):
        results = search_files("keyword", self.test_dir)
        self.assertEqual(len(results), 1)
        self.assertEqual(results[0][0], os.path.join(self.test_dir, 'test_file.txt'))

    def test_search_files_excel(self):
        results = search_files("keyword", self.test_dir)
        self.assertEqual(len(results), 1)
        self.assertEqual(results[0][0], os.path.join(self.test_dir, 'test_file.xlsx'))

    def test_format_results(self):
        captured_output = StringIO()
        sys.stdout = captured_output
        format_results([("file_path", "sheet_name", "coordinate", [("keyword", 1, 2)])])
        sys.stdout = sys.__stdout__  # Zurücksetzen des sys.stdout
        expected_output = "Ergebnisse:\nDatei: file_path\n  Info: sheet_name - Übereinstimmung: 'keyword'\n"
        self.assertEqual(captured_output.getvalue(), expected_output)


if __name__ == '__main__':
    unittest.main()
