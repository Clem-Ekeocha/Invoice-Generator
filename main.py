"""
The step-by-step process to converting excel sales files into Invoices
"""

import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

# Using the glob module to import the filepaths of the xlsx files into a list
filepaths = glob.glob("invoices/*.xlsx")

# Iterate over each filepath,add them to a df dataframe & create pdf pages
for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()

    # Using the Path import to extract just the name of the file in the path
    filename = Path(filepath).stem

    # Split the file name at the first occurence of '-' into 2 lists and extract the first member
    invoice_nr = filename.split('-')[0]
    pdf.set_font(family='Times', size=16, style='B')

    # The extracted first member is used at the invoice number
    pdf.cell(w=50, h=8, txt=f"Invoice nr.{invoice_nr}")
    pdf.output(f"Pdf_Output/{filename}_output.pdf")

