import os
import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path


def generate(invoices_path, pdfs_path, product_id, product_name, amount_purchased, price_per_unit, total_price):
    filepaths = glob.glob(f"{invoices_path}/*.xlsx")
    """
    Generates PDF invoices from Excel files.

    This function reads Excel files from the specified `invoices_path`, extracts invoice data, and generates PDF invoices
    saved to the specified `pdfs_path`. Each Excel file should be named in the format 'invoice_nr-date.xlsx' and contain
    a sheet named 'Sheet 1' with the following columns: product_id, product_name, amount_purchased, price_per_unit, and total_price.

    Parameters:
    invoices_path (str): The directory path where the Excel invoice files are located.
    pdfs_path (str): The directory path where the generated PDF invoices will be saved.
    product_id (str): The column name for the product ID in the Excel file.
    product_name (str): The column name for the product name in the Excel file.
    amount_purchased (str): The column name for the amount purchased in the Excel file.
    price_per_unit (str): The column name for the price per unit in the Excel file.
    total_price (str): The column name for the total price in the Excel file.

    Returns:
    None

    """

    for filepath in filepaths:

        pdf = FPDF(orientation="P", unit="mm", format="A4")
        pdf.add_page()

        filename = Path(filepath).stem
        invoice_nr, date = filename.split("-")

        pdf.set_font(family="Times", size=16, style="B")
        pdf.cell(w=50, h=8, text=f"Invoice nr.{invoice_nr}", ln=1)

        pdf.set_font(family="Times", size=16, style="B")
        pdf.cell(w=50, h=8, text=f"Date: {date}", ln=1)

        df = pd.read_excel(filepath, sheet_name="Sheet 1")

        # Add a header
        columns = df.columns
        columns = [item.replace("_", " ").title() for item in columns]
        pdf.set_font(family="Times", size=10, style="B")
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, text=columns[0], border=1)
        pdf.cell(w=70, h=8, text=columns[1], border=1)
        pdf.cell(w=30, h=8, text=columns[2], border=1)
        pdf.cell(w=30, h=8, text=columns[3], border=1)
        pdf.cell(w=30, h=8, text=columns[4], border=1, ln=1)

        # Add rows to the table
        for index, row in df.iterrows():
            pdf.set_font(family="Times", size=10)
            pdf.set_text_color(80, 80, 80)
            pdf.cell(w=30, h=8, text=str(row[product_id]), border=1)
            pdf.cell(w=70, h=8, text=str(row[product_name]), border=1)
            pdf.cell(w=30, h=8, text=str(row[amount_purchased]), border=1)
            pdf.cell(w=30, h=8, text=str(row[price_per_unit]), border=1)
            pdf.cell(w=30, h=8, text=str(row[total_price]), border=1, ln=1)

        total_sum = df[total_price].sum()
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, text="", border=1)
        pdf.cell(w=70, h=8, text="", border=1)
        pdf.cell(w=30, h=8, text="", border=1)
        pdf.cell(w=30, h=8, text="", border=1)
        pdf.cell(w=30, h=8, text=str(total_sum), border=1, ln=1)

        # Add total sum sentence
        pdf.set_font(family="Times", size=10, style="B")
        pdf.cell(w=30, h=8, text=f"The total price is {total_sum}", ln=1)

        # Add company name and logo
        pdf.set_font(family="Times", size=14, style="B")
        pdf.cell(w=25, h=8, text=f"PythonHow")
        pdf.image("pythonhow.png", w=10)

        os.makedirs(pdfs_path, exist_ok=True)
        pdf.output(f"{pdfs_path}/{filename}.pdf")
