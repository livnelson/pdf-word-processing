import os
from PyPDF2 import PdfReader, PdfWriter

def split_pdf(input_path, output_dir, pages_per_pdf):
    # Check if input file exists
    if not os.path.isfile(input_path):
        raise FileNotFoundError(f"The file at {input_path} does not exist.")
    
    # Check if output directory exists
    if not os.path.isdir(output_dir):
        raise FileNotFoundError(f"The directory at {output_dir} does not exist.")
    
    input_pdf = PdfReader(input_path)
    total_pages = len(input_pdf.pages)
    total_pdfs = total_pages // pages_per_pdf + (1 if total_pages % pages_per_pdf != 0 else 0)

    for i in range(total_pdfs):
        start_page = i * pages_per_pdf
        end_page = min((i + 1) * pages_per_pdf, total_pages)

        output_pdf = PdfWriter()
        for page_num in range(start_page, end_page):
            output_pdf.add_page(input_pdf.pages[page_num])

        output_filename = f'insert-file-name_{i + 1}.pdf'
        output_path = os.path.join(output_dir, output_filename)
        with open(output_path, 'wb') as output_file:
            output_pdf.write(output_file)

#! Update the file path before running the script
if __name__ == '__main__':
    input_path = '[insert-file-path-here]'  # Specify the path to your input PDF file
    output_dir = '[insert-file-path-here]'  # Specify the directory to save the output PDFs
    pages_per_pdf = 2  # Number of pages per output PDF

    split_pdf(input_path, output_dir, pages_per_pdf)