import glob
import pandas as pd
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    pdf = FPDF(orientation="P", unit="mm", format="legal")
    pdf.add_page()
    filename = Path(filepath).stem
    invoice_number = filename.split("-")[0]
    invoice_date = filename.split("-")[1]

    pdf.set_font("Times", size=18)
    pdf.cell(h=8, w=50,  txt=f"Invoice #: {invoice_number}", ln=1)
    pdf.cell(h=8, w=50, txt=f"Date: {invoice_date}", ln=1)

    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # Add a header
    headers = df.columns
    headers = [item.replace('_', ' ').title() for item in headers]
    pdf.set_font(family="Times", size=10, style="B")
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=headers[0], border=1)
    pdf.cell(w=70, h=8, txt=headers[1], border=1)
    pdf.cell(w=40, h=8, txt=headers[2], border=1)
    pdf.cell(w=30, h=8, txt=headers[3], border=1)
    pdf.cell(w=20, h=8, txt=headers[4], ln=1, border=1)

    # Add invoice data
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row['product_id']), border=1)
        pdf.cell(w=70, h=8, txt=str(row['product_name']), border=1)
        pdf.cell(w=40, h=8, txt=str(row['amount_purchased']), border=1)
        pdf.cell(w=30, h=8, txt=str(row['price_per_unit']), border=1)
        pdf.cell(w=20, h=8, txt=str(row['total_price']), ln=1, border=1)

    pdf.output(f"PDFs/{filename}.pdf")
