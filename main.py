import pandas as pd
import glob

filepaths = glob.glob("excel_files/*xlsx")

for filepath in filepaths:
    excel_data = pd.read_excel(filepath, sheet_name="Sheet 1")
    print(excel_data)
