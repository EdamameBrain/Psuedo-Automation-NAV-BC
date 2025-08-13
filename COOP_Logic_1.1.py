import pandas as pd
import re

def read_txt_file(file_path):
    """
    Reads and parses a .txt file with Customer Number, Invoice Number, and Amount.
    Returns a list of dictionaries.
    """
    records = []
    with open(file_path, 'r') as file:
        for line in file:
            parts = re.split(r'\s{2,}|\t', line.strip())
            if len(parts) == 3:
                customer_no, invoice_no, amount_str = parts
                amount = float(re.sub(r'[^\d.-]', '', amount_str))
                records.append({
                    "CustomerNumber": customer_no,
                    "InvoiceNumber": invoice_no,
                    "Amount": amount
                })
    return records

def generate_ascr_numbers(start_ascr, count):
    """
    Generates sequential ASCR numbers from a starting ASCR string like 'ASCR-123456'.
    """
    prefix, start_num = start_ascr.split('-')
    return [f"{prefix}-{int(start_num) + i:06d}" for i in range(count)]

def populate_sales_header(records, ascr_numbers):
    """
    Populates the Sales Header tab as a DataFrame.
    Applies blanking rules and label changes for non-SI entries.
    """
    header_data = []
    for record, ascr in zip(records, ascr_numbers):
        invoice_no = record["InvoiceNumber"]
        is_si = invoice_no.upper().startswith("SI")
        header_data.append({
            "Document Type": "Credit Memo",
            "No.": ascr,
            "Sell-to Customer No.": record["CustomerNumber"],
            "Bill-to Customer No.": record["CustomerNumber"],
            "Posting Description": f"Credit Memo {ascr}",
            "Location Code": "W01",
            "Applies-to Doc. Type": "Invoice" if is_si else "",
            "Applies-to Doc. No.": invoice_no if is_si else "",
            "External Document No.": f"COOP {invoice_no}" if not is_si else f"COOP TO COVER {invoice_no}"
        })
    return pd.DataFrame(header_data)

def populate_sales_line(records, ascr_numbers):
    """
    Populates the Sales Line tab as a DataFrame with two lines per ASCR.
    Applies label shortening for non-SI entries and integer line numbers.
    """
    line_data = []
    for record, ascr in zip(records, ascr_numbers):
        invoice_no = record["InvoiceNumber"]
        is_si = invoice_no.upper().startswith("SI")
        # Line 1
        line_data.append({
            "Document Type": "Credit Memo",
            "Document No.": ascr,
            "Line No.": 10000,
            "Type": "G/L Account",
            "No.": 226000,
            "Location Code": "W01",
            "Description": "ACCR-COOP ADVERTISING",
            "Quantity": 1,
            "Unit Price": record["Amount"],
            "Amount": record["Amount"],
            "Tax Group Code": "NONTAXABLE"
        })
        # Line 2
        line_data.append({
            "Document Type": "Credit Memo",
            "Document No.": ascr,
            "Line No.": 20000,
            "Type": "G/L Account",
            "No.": 226000,
            "Location Code": "W01",
            "Description": f"COOP {invoice_no}" if not is_si else f"COOP TO COVER {invoice_no}",
            "Quantity": 0,
            "Unit Price": 0,
            "Amount": 0,
            "Tax Group Code": "NONTAXABLE"
        })
    return pd.DataFrame(line_data)

def generate_credit_memo_excel(txt_file_path, start_ascr, output_path):
    """
    Main function to generate the Excel file for COOP Credit Memo batch.
    """
    records = read_txt_file(txt_file_path)
    ascr_numbers = generate_ascr_numbers(start_ascr, len(records))
    header_df = populate_sales_header(records, ascr_numbers)
    line_df = populate_sales_line(records, ascr_numbers)

    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        header_df.to_excel(writer, sheet_name='Sales Header', index=False)
        line_df.to_excel(writer, sheet_name='Sales Line', index=False)
    print(f"Credit memo Excel generated: {output_path}")
