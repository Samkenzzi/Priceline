import streamlit as st
import pandas as pd
from fpdf import FPDF
from datetime import datetime
import os
from collections import defaultdict

def generate_packing_slips_from_excel(file, file_name):
    xls = pd.ExcelFile(file)
    df = xls.parse("Sheet1")
    items_df = df[df["Item Description"].notna() & df["Quantity"].notna()]
    if items_df.empty:
        return []

    class PackingSlipPDF(FPDF):
        def header(self):
            self.set_font("Arial", "B", 12)
            self.cell(0, 10, "Kingsbury Court PTY LTD (KENZZI)", ln=True, align="L")
            self.set_font("Arial", "B", 16)
            self.cell(0, 10, "Packing Slip", ln=True, align="L")
            self.ln(5)

        def shipping_info_and_address(self, meta):
            order_date = pd.to_datetime(meta["Date of Order"]).strftime("%d/%m/%Y")
            order_number = str(int(meta["Sales Order Number"]))
            phone = str(int(meta["Phone"]))
            ship_to_lines = [
                meta["Ship_Addressee"],
                meta["Ship_Address Line 1"],
                f"{meta['Ship_City']}, {meta['Ship_State']},{int(meta['Ship_Postcode'])}"
            ]

            self.set_font("Arial", "", 12)
            self.cell(100, 10, "SHIP TO", ln=False)
            self.cell(0, 10, f"Recipient Order Date {order_date}", ln=True)
            for line in ship_to_lines:
                self.cell(100, 10, line, ln=False)
                if line == ship_to_lines[0]:
                    self.cell(0, 10, f"Order Number {order_number}", ln=True)
                elif line == ship_to_lines[1]:
                    self.cell(0, 10, f"Phone {phone}", ln=True)
                elif line == ship_to_lines[2]:
                    self.cell(0, 10, f"Purchase Order {order_number}", ln=True)
                else:
                    self.ln()
            self.ln(5)

        def items_table(self, data):
            self.set_font("Arial", "B", 12)
            self.cell(40, 10, "Product Code", border=1)
            self.cell(80, 10, "Description", border=1)
            self.cell(40, 10, "Item_Code", border=1)
            self.cell(30, 10, "Quantity", border=1)
            self.ln()
            self.set_font("Arial", "", 12)
            for _, row in data.iterrows():
                self.cell(40, 10, str(int(row["Customer No#"])), border=1)
                self.cell(80, 10, row["Item Description"], border=1)
                self.cell(40, 10, str(int(row["Item_Code"])), border=1)
                self.cell(30, 10, str(int(row["Quantity"])), border=1)
                self.ln()

    output_dir = r"C:\\Users\\suvid\\Downloads"
    os.makedirs(output_dir, exist_ok=True)

    pdf_paths = []
    grouped = items_df.groupby("Sales Order Number")
    for order_number, group in grouped:
        pdf = PackingSlipPDF()
        pdf.add_page()
        meta = group.iloc[0]
        pdf.shipping_info_and_address(meta)
        pdf.items_table(group)
        pdf.cell(0, 10, "Packing Slip", ln=True)

        output_path = os.path.join(output_dir, f"{file_name}_PO_{int(order_number)}_packing_slip.pdf")
        pdf.output(output_path)
        pdf_paths.append(output_path)

    return pdf_paths

# Streamlit interface
st.title("Packing Slip Generator")
st.write("Upload one or more Priceline Kenzzi Excel order templates to generate packing slips.")

uploaded_files = st.file_uploader("Upload Excel files", type="xlsx", accept_multiple_files=True)

if uploaded_files:
    for uploaded_file in uploaded_files:
        with st.spinner(f"Generating packing slip(s) for {uploaded_file.name}..."):
            result_pdf_paths = generate_packing_slips_from_excel(uploaded_file, os.path.splitext(uploaded_file.name)[0])
            if result_pdf_paths:
                for path in result_pdf_paths:
                    with open(path, "rb") as f:
                        st.download_button(
                            label=f"Download {os.path.basename(path)}",
                            data=f,
                            file_name=os.path.basename(path),
                            mime="application/pdf"
                        )
            else:
                st.error(f"No valid item data found in {uploaded_file.name}.")
