import os
from PyPDF2 import PdfReader, PdfWriter

def split_pdf(input_path, output_dir, pages_per_pdf):
    input_pdf = PdfReader(input_path)

    total_pages = len(input_pdf.pages)
    total_pdfs = total_pages // pages_per_pdf + (1 if total_pages % pages_per_pdf != 0 else 0)

    for i in range(total_pdfs):
        start_page = i * pages_per_pdf
        end_page = min((i + 1) * pages_per_pdf, total_pages)

        output_pdf = PdfWriter()
        for page_num in range(start_page, end_page):
            output_pdf.add_page(input_pdf.pages[page_num])

        output_filename = f'energy_community_{i + 1}.pdf'
        output_path = os.path.join(output_dir, output_filename)
        with open(output_path, 'wb') as output_file:
            output_pdf.write(output_file)

if __name__ == '__main__':
    input_path = '/Users/livnelson/Downloads/merged.pdf'  # Specify the path to your input PDF file
    output_dir = '/Users/livnelson/Downloads/NEI_DOCS'  # Specify the directory to save the output PDFs
    pages_per_pdf = 2  # Number of pages per output PDF

    split_pdf(input_path, output_dir, pages_per_pdf)
    split_pdf(input_path, output_dir, pages_per_pdf)                                                        




# !Path to  Word template
# /Users/livnelson/Downloads/XXX_Energy_Community _Template.docx

# !Path to  PDF file
# /Users/livnelson/Downloads/ALL_DOCS.pdf

# !Path to Excel data file
# /Users/livnelson/Downloads/Combined_Leads_Owner_Opperated.xlsx

# !Path to outfut folder
# /Users/livnelson/Downloads/NEI_DOCS