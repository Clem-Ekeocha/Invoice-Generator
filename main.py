"""
The step-by-step process to converting excel sales files into Invoices
"""

import pandas as pd
import glob

# Using the glob module to import the filepaths of the xlsx files into a list
filepaths = glob.glob("invoices/*.xlsx")

# Iterate over each filepath to add them to the df dataframe
for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    print(df)