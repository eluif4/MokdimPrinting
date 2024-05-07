# printerconfig.py
import os
from dotenv import load_dotenv
import win32print

# get printer name from .env file
printer_name = os.getenv("PRINTER_NAME")

def get_printer():
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
