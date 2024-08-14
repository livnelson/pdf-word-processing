# PDF and Word Document Processing

This repository contains a set of Python scripts for handling various document-related tasks, including:

1. **Splitting PDFs into smaller files**
2. **Performing a mail merge with a Word template and an Excel file**
3. **Renaming PDF files based on an Excel spreadsheet**

## Table of Contents

1. [Splitting PDFs](#splitting-pdfs)
2. [Mail Merge](#mail-merge)
3. [Renaming PDF Files](#renaming-pdf-files)
4. [Installation](#installation)
5. [Contributing](#contributing)
6. [License](#license)



## Splitting PDFs

**`split_pdf.py`** - This script splits a PDF file into smaller files with a specified number of pages per file.

### Usage

1. Update the file paths in the script:

   ```
   input_path = '[insert-path-to-pdf-file]'
   output_dir = '[insert-output-directory]'
   pages_per_pdf = 2
   ```
2. Run the script:

   ```
   python3 split_pdf.py
   ```



## Mail Merge
**`mail_merge_script.py`** - This script performs a mail merge operation by replacing placeholders in a Word template with data from an Excel spreadsheet.

1. Update the file paths in the script:

   ```
   template_path = '[insert-path-to-word-template]'
   excel_path = '[insert-path-to-excel-file]'
   output_dir = '[insert-output-directory]'
   ```
   
2. Run the script:

   ```
   python3 mail_merge_script.py
   ```
   This command will generate merged Word documents for each row in the Excel file, replacing placeholders in the Word template.

## Renaming PDF Files
**`rename_pdf_files.py`** - This script renames and moves PDF files based on new names provided in an Excel spreadsheet.

1. Update the file paths and parameters in the script:

   ```
   folder_path = '[insert-path-to-folder-containing-pdfs]'
   excel_path = '[insert-path-to-excel-file]'
   output_folder = '[insert-output-folder]'
   start_index = 1
   ```
   
2. Run the script:

   ```
   python3 rename_pdf_files.py
   ```
   This command will rename and move PDF files from the specified folder to the output folder, using names from the Excel file.



## Installation
To run these scripts, ensure you have Python and the necessary libraries installed. You can install the required libraries using:

```
pip install pandas python-docx PyPDF2
```

## License
This project is licensed under the MIT License - see the LICENSE file for details.
