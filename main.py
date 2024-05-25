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

    # The extracted first member is used at the invoice number
    pdf.set_font(family='Times', size=16, style='B')
    pdf.cell(w=50, h=8, txt=f"Invoice nr.{invoice_nr}", ln=1)

    # After the split, the second member of the split is the date
    date = filename.split('-')[1]

    # The extracted second member is used at the date
    pdf.set_font(family='Times', size=16, style='B')
    pdf.cell(w=50, h=8, txt=f"Date nr.{date}", ln=1)

    # Next: Read the table data into the pdf starting with the Headers.
    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    column_headers = df.columns
    column_headers = [item.replace("_", " ").title() for item in column_headers]

    # Define the position and title headers of the table to be created in the pdf
    pdf.set_font(family="Times", size=10, style='B')
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=column_headers[0], border=1)
    pdf.cell(w=60, h=8, txt=column_headers[1], border=1)
    pdf.cell(w=40, h=8, txt=column_headers[2], border=1)
    pdf.cell(w=30, h=8, txt=column_headers[3], border=1)
    pdf.cell(w=30, h=8, txt=column_headers[4], border=1, ln=1)

    # Read the excel file to add rows to created above the table
    # by iterating over the rows of the table in the xlsx.
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8,txt=str(row["product_id"]), border=1)
        pdf.cell(w=60, h=8,txt=str(row["product_name"]), border=1)
        pdf.cell(w=40, h=8,txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8,txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8,txt=str(row["total_price"]), border=1, ln=1)


    pdf.output(f"Pdf_Output/{filename}_output.pdf")

