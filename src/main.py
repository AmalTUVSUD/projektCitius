import win32com.client as win32
from pathlib import Path
import os

def get_file():
    while True:
        file_path = input("Enter the path to the .doc file: ").strip('"\'')

        #Validate file path
        path = Path(file_path)
        if not path.exists():
            print(f"File not found: {file_path}")
            continue
        if path.suffix.lower() != '.doc':
            print("File is not a .doc file")
            continue
        return path