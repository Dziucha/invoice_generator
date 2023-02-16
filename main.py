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
             align="L", ln=1)
    pdf.cell(w=0, h=10, txt=f"Date {date_of_invoice}",
             align="L", ln=1)

    excel_data = pd.read_excel(filepath, sheet_name="Sheet 1")

    pdf.set_font(family="Times", size=10, style="B")
    pdf.set_text_color(80, 80, 80)

    headers = list(excel_data.columns)
    headers = [header.replace("_", " ").title() for header in headers]

    pdf.cell(w=30, h=8, txt=headers[0], align="L", border=1)
    pdf.cell(w=67, h=8, txt=headers[1], align="L", border=1)
    pdf.cell(w=33, h=8, txt=headers[2], align="L", border=1)
    pdf.cell(w=30, h=8, txt=headers[3], align="L", border=1)
    pdf.cell(w=30, h=8, txt=headers[4], align="L", border=1, ln=1)

    for index, row in excel_data.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]),
                 align="L", border=1)
        pdf.cell(w=67, h=8, txt=row["product_name"],
                 align="L", border=1)
        pdf.cell(w=33, h=8, txt=str(row["amount_purchased"]),
                 align="L", border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]),
                 align="L", border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]),
                 align="L", border=1, ln=1)

    total_price = excel_data["total_price"].sum()

    pdf.cell(w=30, h=8, txt="", align="L", border=1)
    pdf.cell(w=67, h=8, txt="", align="L", border=1)
    pdf.cell(w=33, h=8, txt="", align="L", border=1)
    pdf.cell(w=30, h=8, txt="", align="L", border=1)
    pdf.cell(w=30, h=8, txt=f"Total {total_price}",
             align="L", border=1, ln=1)

    pdf.set_font(family="Times", size=10, style="B")
    pdf.set_text_color(0, 0, 0)
    pdf.cell(w=0, h=8, txt=f"The total due amount is: {total_price} Euros.",
             align="L", ln=1)
    pdf.cell(w=0, h=8, txt="PythonHow", align="L", ln=1)

    pdf.output(f"pdf_invoices/{file_name}.pdf")
