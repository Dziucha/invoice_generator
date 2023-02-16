import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("excel_files/*xlsx")

for filepath in filepaths:
    file_name = Path(filepath).stem

    pdf = FPDF(orientation="P", format="A4", unit="mm")
    pdf.set_auto_page_break(auto=False, margin=0)

    pdf.add_page()
    pdf.set_font(family="Times", size=16, style="B")

    invoice_nr, date_of_invoice = file_name.split(sep="-")

    pdf.cell(w=0, h=10, txt=f"Invoice nr. {invoice_nr}",
             align="L", border=1, ln=1)
    pdf.cell(w=0, h=10, txt=f"Date {date_of_invoice}",
             align="L", border=1, ln=1)

    excel_data = pd.read_excel(filepath, sheet_name="Sheet 1")


    total_price = excel_data["total_price"].sum()

    pdf.set_font(family="Times", size=10, style="B")
    pdf.cell(w=0, h=8, txt=f"The total due amount is: {total_price} Euros.",
             align="L", border=1, ln=1)
    pdf.cell(w=0, h=8, txt="PythonHow", align="L", border=1, ln=1)

    pdf.output(f"pdf_invoices/{file_name}.pdf")
