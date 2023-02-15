import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("excel_files/*xlsx")

for filepath in filepaths:
    excel_data = pd.read_excel(filepath, sheet_name="Sheet 1")

    file_name = Path(filepath).stem

    pdf = FPDF(orientation="P", format="A4", unit="mm")
    pdf.set_auto_page_break(auto=False, margin=0)

    pdf.add_page()
    pdf.set_font(family="Times", size=16, style="B")

    file_name_data = file_name.split(sep="-")
    invoice_nr = file_name_data[0]
    pdf.cell(w=0, h=10, txt=f"Invoice nr. {invoice_nr}",
             align="L", border=1, ln=1)
    date_of_invoice = file_name_data[1]
    pdf.cell(w=0, h=10, txt=f"Date {date_of_invoice}",
             align="L", border=1, ln=1)

    total_price = 0
    for index, row in excel_data.iterrows():
        total_price = total_price + float(row["total_price"])

    pdf.set_font(family="Times", size=10, style="B")
    pdf.cell(w=0, h=8, txt=f"The total due amount is: {total_price} Euros.",
             align="L", border=1, ln=1)
    pdf.cell(w=0, h=8, txt="PythonHow", align="L", border=1, ln=1)

    pdf.output(f"pdf_invoices/{file_name}.pdf")
