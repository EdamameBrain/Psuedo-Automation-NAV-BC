
from openpyxl import Workbook

def export_to_excel_with_sales_invoice_format_and_note(audit_trails, ra_number, invoice_no, customer_1, customer_2, output_path=None):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    # First block: Original journal entry format
    journal_headers = [
        'Type', 'No.', 'Description', 'Location Code', 'Quantity', 'Unit of Measure Code',
        'Unit Price Excl. Tax', 'Tax Group Code', 'Line Amount Excl. Tax', 'Amount Including Tax',
        'Line Discount %', 'Qty. to Assign', 'Qty. Assigned'
    ]
    ws.append(journal_headers)

    for invoice, item_no, qty, price in audit_trails:
        try:
            qty_float = float(qty)
            price_float = float(price)
            line_amount = qty_float * price_float
        except (TypeError, ValueError):
            qty_float = 0
            price_float = 0
            line_amount = 0

        unit_price_fmt = f"{price_float:.2f}"
        line_amount_fmt = f"{line_amount:.2f}"

        ws.append([
            "G/L Account", 485300, "MISC. ALLOWANCES", "W01", qty_float, "EA",
            unit_price_fmt, "NONTAXABLE", line_amount_fmt, line_amount_fmt,
            " ", 0, " "
        ])
        description = f"({qty}) {item_no} @ ${unit_price_fmt} EA {invoice}"
        ws.append([
            " ", " ", description, " ", " ", " ",
            " ", " ", " ", 0,
            " ", 0, " "
        ])

    rebill_note = f"REBILL {invoice_no} {customer_1} TO {customer_2}"
    ws.append([" ", " ", rebill_note, " ", " ", " ", " ", " ", " ", 0, " ", 0, " "])

    # Spacer row before the second format
    ws.append([])

    # Second block: Sales Invoice format
    sales_headers = [
        "Type", "No.", "Description", "Location Code", "Quantity", "Unit of Measure Code",
        "Unit Price Excl. Tax", "Tax Group Code", "Line Discount %", "Line Amount Excl. Tax",
        "Amount Including Tax", "Qty. to Assign"
    ]
    ws.append(sales_headers)

    for invoice, item_no, qty, price in audit_trails:
        try:
            qty_float = float(qty)
            price_float = float(price)
            line_amount = qty_float * price_float
        except (TypeError, ValueError):
            qty_float = 0
            price_float = 0
            line_amount = 0

        unit_price_fmt = f"{price_float:.2f}"
        line_amount_fmt = f"{line_amount:.2f}"
        description = f"({qty}) {item_no} @ ${unit_price_fmt} EA {invoice}"

        ws.append([
            "G/L Account", 485300, description, "W01", qty_float, "EA",
            unit_price_fmt, "", "", line_amount_fmt, line_amount_fmt, 0
        ])

    ws.append([" ", " ", rebill_note, " ", " ", " ", " ", " ", " ", " ", " ", 0])

    if not output_path:
        output_path = f"Rebill {invoice_no} {customer_1} to {customer_2}.xlsx"
    elif not output_path.lower().endswith(".xlsx"):
        output_path = f"Rebill {invoice_no} {customer_1} to {customer_2}.xlsx"

    wb.save(output_path)
    return output_path
