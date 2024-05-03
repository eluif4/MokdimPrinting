# printerconfig.py

import win32print

def get_printer():
    printer_name = "HP97EA39 (HP OfficeJet Pro 8020 series)"  # Replace with your printer name
    try:
        printer_handler = win32print.OpenPrinter(printer_name)
        return printer_handler
    except Exception as e:
        print(f"Error opening printer '{printer_name}': {e}")
        return None

# Usage example:
# printer_handler = get_printer()
# if printer_handler:
#     # Use the printer for printing
#     # ...
