import win32com.client as win32
from pathlib import Path
import os


class DocHandler:
    def __init__(self):
        self.word = win32.Dispatch("Word.Application")
        self.word.Visible = False
        self.doc = None

    def open_doc(self, file_path):
        if not Path(file_path).exists():
            raise FileNotFoundError(f"File not found: {input_path}")
        if not file_path.lower().endswith('.doc'):
            raise ValueError("File is not a .doc file")
        return self.word.Documents.Open(os.path.abspath(file_path))


    def close_doc(self):
        if self.doc:
            self.doc.Close(SaveChanges=False)
            self.doc = None

    def quit_word(self):
        self.word.Quit()

    def get_tables(self):
        if not self.doc:
            raise ValueError("No document is open")
        tables = self.doc.Tables
        return tables
    
    def save_doc(self, file_path):
        if not self.doc:
            raise ValueError("No document is open")
        self.doc.SaveAs(os.path.abspath(file_path))
        self.doc.Close(SaveChanges=False)


# Initialize
doc_handler = DocHandler()

# Open file
doc = doc_handler.open_doc("input/Test_TRF.doc")

# Process tables
tables = doc_handler.get_tables(doc)
for table in tables:
    # Modify table content (example: uppercase first cell)
    table.Cell(1, 1).Range.Text = table.Cell(1, 1).Range.Text.upper()

# Save changes
doc_handler.save_as_doc(doc, "output/Modified_TRF.doc")

# Clean up
doc_handler.close()