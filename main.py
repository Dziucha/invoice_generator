import pandas as pd
import glob
from fpdf import FPDF

filepaths = glob.glob("excel_files/*xlsx")

for filepath in filepaths:
    excel_data = pd.read_excel(filepath, sheet_name="Sheet 1")
    pdf = FPDF(orientation="P", format="A4", unit="mm")

    file_name = filepath.strip(".xlsx")
    file_name = file_name.lstrip("excel_files/")
    pdf.output(f"pdf_invoices/{file_name}.pdf")
