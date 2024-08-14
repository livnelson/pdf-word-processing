import os
import pandas as pd
import fnmatch # Import fnmatch module for file pattern matching
import re # Import re module for regular expressions

def sanitize_file_name(name):
    # Remove invalid characters from the file name
    sanitized_name = re.sub(r'[<>:"/\\|?*]', '', name)
    return sanitized_name

def get_unique_file_name(file_name, output_folder):
    if os.path.exists(os.path.join(output_folder, file_name)):  # Check if the file already exists in the output folder
        base_name, extension = os.path.splitext(file_name)  # If the file exists, add a numerical suffix to make it unique
        suffix = 1
        new_file_name = f"{base_name}_{suffix}{extension}"
        while os.path.exists(os.path.join(output_folder, new_file_name)):
            suffix += 1
            new_file_name = f"{base_name}_{suffix}{extension}"
        return new_file_name
    else:
        return file_name

def rename_pdf_files(folder_path, excel_path, output_folder, start_index):
    # Load the Excel spreadsheet into a pandas DataFrame
    df = pd.read_excel(excel_path)
    
    # Iterate over the rows in the DataFrame starting from the specified index
    for index, row in df.iloc[start_index-1:].iterrows():
        # Extract the new file name from the first column
        new_file_name = row.iloc[0]  # Assumes the first column contains the new file names
        # Sanitize the file name
        new_file_name = sanitize_file_name(new_file_name)
        # Get a unique file name for the output folder
        new_file_name = get_unique_file_name(new_file_name + ".pdf", output_folder)
        
        # Construct the current file name pattern to match in the folder
        current_file_pattern = f"*_{index + 1}.pdf"  # Assuming index starts from 0
        
        # Find the matching PDF file in the folder
        matching_files = [file for file in os.listdir(folder_path) if fnmatch.fnmatch(file, current_file_pattern)]
        if matching_files:
            # Rename the first matching file to the new file name
            old_file_path = os.path.join(folder_path, matching_files[0])
            new_file_path = os.path.join(output_folder, new_file_name)
            
            # Move the renamed file to the output folder
            os.rename(old_file_path, new_file_path)
            print(f"Renamed {matching_files[0]} to {new_file_name} and moved to {output_folder}")
        else:
            print(f"No matching file found for row {index + 1} in the specified folder.")

#! Update the file path before running the script
if __name__ == "__main__":
    folder_path = "[insert-path-here]"  # Specify the folder containing the PDF files
    excel_path = "[insert-path-here]"  # Specify the path to Excel spreadsheet
    output_folder = "[insert-path-here]"  # Specify the output folder for renamed files
    start_index = 1  # Specify the index to start from

    rename_pdf_files(folder_path, excel_path, output_folder, start_index)