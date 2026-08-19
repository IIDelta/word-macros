import os
import sys
import unittest
import time

try:
    import win32com.client
except ImportError:
    print("Error: pywin32 is not installed. Please run 'pip install -r requirements.txt'")
    sys.exit(1)

class TestWordMacros(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        cls.dotm_path = os.path.join(cls.base_dir, 'dist', 'MedicalWritingTools.dotm')
        cls.test_doc_path = os.path.join(cls.base_dir, 'tests', 'mock_data.docx')
        
        if not os.path.exists(cls.dotm_path):
            raise FileNotFoundError(f"Cannot find {cls.dotm_path}. Please run build.py first.")
            
        print("Starting Word for testing...")
        try:
            cls.word = win32com.client.Dispatch("Word.Application")
            cls.word.Visible = False
            cls.word.AutomationSecurity = 1 # msoAutomationSecurityLow
        except Exception as e:
            print(f"Failed to start Word: {e}")
            sys.exit(1)

    @classmethod
    def tearDownClass(cls):
        cls.word.Quit()

    def setUp(self):
        # Create a fresh mock document for each test
        self.doc = self.word.Documents.Add(Template=self.dotm_path)
        
    def tearDown(self):
        self.doc.Close(SaveChanges=False)

    def test_remove_multiple_spaces(self):
        # Insert test data
        self.word.Selection.TypeText("This  has   multiple    spaces.")
        
        # Run macro
        self.word.Application.Run("MW_RemoveMultipleSpaces")
        
        # Verify
        content = self.doc.Content.Text
        self.assertNotIn("  ", content)
        self.assertIn("This has multiple spaces.", content)

    def test_yellow_redline(self):
        # Insert test data with tracking
        self.doc.TrackRevisions = True
        self.word.Selection.TypeText("Inserted Text.")
        
        # Run macro
        self.word.Application.Run("MW_YellowHighlightRedline")
        
        # Verify track changes were accepted/rejected
        self.assertEqual(self.doc.Revisions.Count, 0)
        
        # Check if highlighting was applied (wdYellow = 7)
        # We find the word 'Inserted'
        rng = self.doc.Content
        rng.Find.Text = "Inserted"
        rng.Find.Execute()
        self.assertEqual(rng.HighlightColorIndex, 7)

if __name__ == '__main__':
    unittest.main()
