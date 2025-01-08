import pandas as pd
import glob
from fpdf import FPDF
import time
from pathlib import Path

filepaths = glob.glob("invoices/*xlsx")
for filepath in filepaths:
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.set_auto_page_break(auto=False, margin=0)
    pdf.add_page()

    filename = Path(filepath).stem
    invoice_nr = filename.split("-")[0]
    date = filename.split("-")[1]

    pdf.set_font('Times', 'B', 16)
    pdf.cell(0, 10, f'Invoice Number.{invoice_nr}', align='L', ln=1)
    pdf.cell(0, 10, f'Date {time.strftime('%Y.%m.%d')}', align='L', ln=2)
    pdf.cell(0, 10, f'Date {date}', align='L', ln=2)

    df = pd.read_excel(filepath, sheet_name='Sheet 1')
    columns = list(df.columns)
    columns = [column.replace("_", " ").title() for column in columns]
    pdf.set_font('Times', 'B', 10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(30, 8, columns[0], border=1)
    pdf.cell(60, 8, columns[1], border=1)
    pdf.cell(40, 8, columns[2], border=1)
    pdf.cell(30, 8, columns[3], border=1)
    pdf.cell(30, 8, columns[4], border=1, ln=1)
    total_price = 0
    for index, row in df.iterrows():
        pdf.set_font('Times', 'B', 10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(30, 8, str(row["product_id"]), border=1)
        pdf.cell(60, 8, str(row["product_name"]), border=1)
        pdf.cell(40, 8, str(row["amount_purchased"]), border=1)
        pdf.cell(30, 8, str(row["price_per_unit"]), border=1)
        pdf.cell(30, 8, str(row["total_price"]), border=1, ln=1)
        total_price = total_price+row["total_price"]
    pdf.set_font('Times', 'B', 10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(30, 8, "", border=1)
    pdf.cell(60, 8, "", border=1)
    pdf.cell(40, 8, "", border=1)
    pdf.cell(30, 8, "", border=1)
    pdf.cell(30, 8, str(total_price), border=1, ln=1)

    pdf.set_font('Times', 'B', 14)
    pdf.cell(30, 8, f"The Total Price is {total_price}", ln=1)
    print(total_price)
    pdf.output(f"PDFs/{filename}.pdf")
