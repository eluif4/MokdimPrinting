from pathlib import Path
import win32com.client
import win32print
import win32ui

# Printer config
# When converting to prod, connect to the printer once at the beginning and then run the scraping and printing code
printer_name = "HP LaserJet MFP M426fdn (5F550F)"
printer_handler = win32print.OpenPrinter(printer_name)
default_printer_info = win32print.GetPrinter(printer_handler, 2)
printer_dc = win32ui.CreateDC()
printer = printer_dc.CreatePrinterDC(printer_name)

def get_printer():
    return printer