File Hash Generator

This Python application provides a graphical user interface for generating file hashes (MD5, SHA1, and SHA256) for all files within a specified directory. It also facilitates the conversion of the generated data into Excel and PDF formats for easy storage and sharing.

Features

  - Graphical User Interface: A user-friendly interface built with tkinter for easy interaction.
  - Hashing Algorithms: Supports three common hashing algorithms - MD5, SHA1, and SHA256.
  - Excel Export: Saves the hash data to an Excel spreadsheet for easy management and analysis.
  - PDF Conversion: Converts the Excel file into a well-formatted PDF document for sharing and printing.

Dependencies

  - tkinter: Used for creating the graphical user interface.
  - os: Provides file and directory handling functions.
  - hashlib: Enables the calculation of hash values.
  - pandas: Used for creating and manipulating data in the form of dataframes.
  - fpdf: Allows for the creation of PDF documents.
  - openpyxl: Used to work with Excel files.

How to Use
  
  - Run the application.
  - Click the "Browse" button to select the directory containing the files you want to hash.
  - Choose the desired hashing algorithms (MD5, SHA1, and SHA256).
  - Click "Generate Hashes" to calculate and save the file hashes.

The generated hashes are stored in an Excel spreadsheet called "Files Hash List.xlsx" and a PDF document in the selected target directory.

Simplify your file hashing tasks and enhance data management with this efficient utility!
