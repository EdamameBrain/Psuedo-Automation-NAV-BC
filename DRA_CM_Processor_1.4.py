
from openpyxl import Workbook

def process_extracted_audit_trails(audit_trails, ra_number):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    headers = [
        'Type', 'No.', 'Description', 'Location Code', 'Quantity', 'Unit of Measure Code',
        'Unit Price Excl. Tax', 'Tax Group Code', 'Line Amount Excl. Tax', 'Amount Including Tax',
        'Line Discount %', 'Qty. to Assign', 'Qty. Assigned'
    ]
    ws.append(headers)

    for invoice, item_no, qty, price in audit_trails:
        try:
            qty_float = float(qty)
            price_float = float(price)
            line_amount = qty_float * price_float
        except (TypeError, ValueError):
            line_amount = 0.00

        unit_price_fmt = f"{price_float:.2f}"
        line_amount_fmt = f"{line_amount:.2f}"

        ws.append([
            "G/L Account", 488000, "DEFECTIVE ALLOWANCES", "W01", qty, "EA",
            unit_price_fmt, "NONTAXABLE", line_amount_fmt, line_amount_fmt,
            " ", 0, " "
        ])
        description = f"({qty}) {item_no} @ ${unit_price_fmt} EA {invoice}"
        ws.append([
            " ", " ", description, " ", " ", " ",
            " ", " ", " ", 0,
            " ", 0, " "
        ])

    ws.append([" ", " ", f"DAMAGES TO {ra_number}", " ", " ", " ", " ", " ", " ", 0, " ", 0, " "])
    ws.append([" ", " ", "DESTROY IN FIELD", " ", " ", " ", " ", " ", " ", 0, " ", 0, " "])

    output_path = f"{ra_number}.xlsx"
    wb.save(output_path)
    return output_path

if __name__ == "__main__":
    # Example usage placeholder for ChatGPT web
    # Upload image, extract audit_trails via GPT, then:
    audit_trails = [
        ("SI397788", "SC100B", 1, 30.00)
    ]
    ra_number = "RA108931"
    process_extracted_audit_trails(audit_trails, ra_number)
