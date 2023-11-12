import glob
from datetime import datetime
from pathlib import Path

import pandas as pd
from fpdf import FPDF

INVOICE_DIRECTORY = "./invoices/"
CELL_WIDTH = 37
CELL_HEIGHT = 12


def format_date(date_string):
    date = datetime.strptime(date_string, "%Y.%m.%d")
    return date.strftime("%d %B %Y")


invoice_files_path = glob.glob(f"{INVOICE_DIRECTORY}*.xlsx")

for file_path in invoice_files_path:
    pdf = FPDF(orientation="P", unit="mm", format="A4")

    pdf.add_page()

    invoice_file_name = Path(file_path).stem
    invoice_num, date = invoice_file_name.split("-")
    formatted_date = format_date(date)

    df = pd.read_excel(Path(file_path))

    pdf.set_auto_page_break(auto=False, margin=10)
    pdf.set_font("Times", size=14, style="B")

    pdf.cell(0, CELL_HEIGHT, txt=f"Invoice Number: {invoice_num}", ln=1, align="L")
    pdf.cell(0, CELL_HEIGHT, txt=f"Date: {formatted_date}", ln=1, align="L")
    pdf.ln()

    pdf.set_left_margin(10)

    pdf.set_font("Times", size=12, style="B")
    for column_name in df.columns:
        column_name = column_name.replace("_", " ").capitalize()
        pdf.cell(CELL_WIDTH, CELL_HEIGHT, txt=str(column_name), border=1, align="C")
    pdf.ln()
    pdf.set_font("Times", size=8)


    for index, row in df.iterrows():
        for column in df.columns:
            value = row[column]
            pdf.cell(CELL_WIDTH, CELL_HEIGHT, txt=str(value), border=1, align="C")
        pdf.ln()  
    cell_width = CELL_WIDTH*4
    total_amount = df["total_price"].sum()
    pdf.cell(cell_width, CELL_HEIGHT, txt="", ln=0, border=1)
    pdf.cell(CELL_WIDTH, CELL_HEIGHT, txt=f"{total_amount}", ln=1, border=1, align="C")
    pdf.ln()

    pdf.set_font("Times", size=12, style="B")
    pdf.cell(0, CELL_HEIGHT, txt=f"Total Amount: {total_amount}", ln=1, align="L")
    pdf.cell(70, CELL_HEIGHT, txt="Thank you for your business!", ln=0, align="L")
    pdf.image("invoice_logo.png", w=12)
    pdf.output(f"{INVOICE_DIRECTORY}{invoice_num}.pdf")