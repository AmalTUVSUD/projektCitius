import win32com.client as win32
from pathlib import Path
import time
import os


class DocHandler:
    def __init__(self):
        self.word = win32.Dispatch("Word.Application")
        self.word.Visible = False
        self.doc = None
        self.file_path = None

    def open_doc(self, file_path):
        if not Path(file_path).exists():
            raise FileNotFoundError(f"File not found! {file_path}")
        if not file_path.lower().endswith('.doc'):
            raise ValueError("File is not a .doc file")
        self.file_path = file_path
        self.doc =self.word.Documents.Open(file_path, ReadOnly=0)
        time.sleep(1)  # Give Word time to fully load the document
        print("Actual type:", type(self.doc))
        print("Has Tables attribute:", hasattr(self.doc, "Tables"))        
        return self.doc

    def close_doc(self):
        if self.doc:
            self.doc.Close()
        self.word.Quit()
    
    def save_doc(self):
        if self.doc is None:
            raise RuntimeError("No document is open")
        self.doc.SaveAs( self.file_path, FileFormat=0)  # 0 for .doc format

    def count_tables(self):
        if self.doc is None:
           raise RuntimeError("No document is open")
        print("Has Tables attribute:", hasattr(self.doc, "Tables"))
        return self.doc.Tables.Count

    def print_first_3_tables(self):
        if self.doc is None:
            raise RuntimeError("No document is open")

        tables = self.doc.Tables
        num_tables = tables.Count
        print(f"Total tables: {num_tables}")

    # Print first 3 tables or all tables if less than 3
        for t_index in range(min(6, num_tables)):
            print(f"\nTable {t_index + 1}:")
            table = tables[t_index + 1]  # Index is 1-based in Word's COM API
            for row in table.Rows:
                row_data = []
                for cell in row.Cells:
                    text = cell.Range.Text.strip().replace('\r\x07', '')  # Clean cell text
                    row_data.append(text)
                print('\t'.join(row_data))