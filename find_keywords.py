# Purpose: 
#   Read pdf files recursively from a folder and compare extracted text 
#   against a list of keywords
#
# Requirements: 
#   python 3.6+ 
#   package dependecies
#       required: PyPDF2
#       optional: pandas, openpyxl
#   folder with pdf files that have selectable text (aka not scanned images)

import os
import pandas as pd
import PyPDF2

keywords = [
    'balloon',
    'dfit',
    'fit',
    'flar',
    'flow',
    'flowing',
    'gais',
    'isip',
    'kick',
    'leak',
    'loss',
    'losing',
    'lost',
    'lot',
    'return',
]


def read_pdf_text(file_path):
    """Read text from all pages in PDF

    Parameters
    ----------
    file_path : str
        Path to pdf file

    Returns
    -------
    str
        Full text from document
    """
    try:
        # Open the file with the context manager
        with open(file_path, 'rb') as pdf_file:
            # Read the file
            read_pdf = PyPDF2.PdfFileReader(pdf_file)
            all_text = ''
            # Loop over each page and add the extractedText to all_text
            for page in read_pdf.pages:
                page_content = page.extractText()
                all_text += str(page_content.encode('utf-8'))
        return all_text
    except PyPDF2.utils.PdfReadError as e:
        print(f"The following file bombed while reading: {file_path}")


def keywords_from_pdfs(pdf_folder_name):
    """Compare the text from a folder of pdf files to a list of keywords

    Parameters
    ----------
    pdf_folder_name : str
        Folder with target pdf files

    Returns
    -------
    tuple  --> keywords, file_counts
        keywords dict of counts for each keyword
        file_counts is a list of dictionaries with counts for each keyword per file
    """
    keywords_dict = {kw: 0 for kw in keywords}
    file_counts = []
    file_path_list = []

    # Use os.walk to get the pdf file paths in the pdf_folder_name
    for path, d, files in os.walk(pdf_folder_name):
        for file in files:
            if os.path.splitext(file.lower())[1] == ".pdf":
                file_path_list.append(os.path.join(path, file))
    
    # Loop through each file
    for fp in file_path_list:
        # Read the text
        text = read_pdf_text(fp)
        if text is None:
            continue
        file_keywords = {'file': fp}
        
        # Loop through the keywords
        for kw in keywords:
            # If a kw in text, count the occurances and add it to the dict
            if kw in text.lower():
                kw_count = text.lower().count(kw)
                keywords_dict[kw] += kw_count
                # Update the big dict of keywords for fun
                file_keywords.update({kw: kw_count})
        
        # Add the file_keywords dict to the list of file_counts
        file_counts.append(file_keywords)
    return keywords_dict, file_counts


if __name__ == "__main__":
    top_level_folder = r'.\scratch\sarah_is_super_cool'
    output_file_name = r'.\output\keyword_counts.xlsx'
    keyword_counts, file_summary = keywords_from_pdfs(top_level_folder)

    # Print out summary count of all keywords
    # print(keyword_counts)

    # Make a dataframe for easy Excel export
    df = pd.DataFrame(file_summary)
    df.to_excel(output_file_name, index=False)
