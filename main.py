import glob
from datetime import datetime
from pathlib import Path

import pandas as pd
from fpdf import FPDF

INVOICE_DIRECTORY = "./invoices/"
CELL_WIDTH = 35
CELL_HEIGHT = 12


def format_date(date_string):
    date = datetime.strptime(date_string, "%Y.%m.%d")
    return date.strftime("%d %B %Y")


invoice_files_path = glob.glob(f"{INVOICE_DIRECTORY}*.xlsx")

for invoice_file_path in invoice_files_path:
    pdf = FPDF(orientation="P", unit="mm", format="A4")

    pdf.add_page()

    invoice_file_name = Path(invoice_file_path).stem
    invoice_num, date = invoice_file_name.split("-")
    formatted_date = format_date(date)

    df = pd.read_excel(Path(invoice_file_path))

    pdf.set_auto_page_break(auto=False, margin=10)
    pdf.set_font("Times", size=8, style="B")
    pdf.set_text_color(150, 150, 150)

    pdf.cell(160, 10, txt=f"Invoice Number: {invoice_num}", ln=0, align="L")
    pdf.cell(0, 10, txt=formatted_date, ln=0, align="L")

    pdf.set_text_color(0, 0, 0)
    pdf.line(10, 18, 200, 18)
    pdf.ln(20)

    pdf.set_left_margin(20)
    pdf.set_font("Times", size=10, style="B")
    pdf.cell(CELL_WIDTH, CELL_HEIGHT, txt=df.columns[0], ln=0, border=1, align="C")
    pdf.cell(CELL_WIDTH, CELL_HEIGHT, txt=df.columns[1], ln=0, border=1, align="C")
    pdf.cell(CELL_WIDTH, CELL_HEIGHT, txt=df.columns[2], ln=0, border=1, align="C")
    pdf.cell(CELL_WIDTH, CELL_HEIGHT, txt=df.columns[3], ln=0, border=1, align="C")
    pdf.cell(CELL_WIDTH, CELL_HEIGHT, txt=df.columns[4], ln=1, border=1, align="C")

    for index, row in df.iterrows():
        pdf.set_font("Times", size=8)
        pdf.cell(CELL_WIDTH, 10, txt=str(row["product_id"]), ln=0, border=1, align="C")
        pdf.cell(
            CELL_WIDTH, 10, txt=str(row["product_name"]), ln=0, border=1, align="C"
        )
        pdf.cell(
            CELL_WIDTH, 10, txt=str(row["amount_purchased"]), ln=0, align="C", border=1
        )
        pdf.cell(
            CELL_WIDTH, 10, txt=str(row["price_per_unit"]), ln=0, align="C", border=1
        )
        pdf.cell(CELL_WIDTH, 10, txt=str(row["total_price"]), ln=1, align="C", border=1)

    pdf.output(f"{INVOICE_DIRECTORY}{invoice_num}.pdf")
