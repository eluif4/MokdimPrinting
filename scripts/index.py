
from pathlib import Path
import win32com.client
import win32print

from printerconfig import get_printer

def print_file(file_path, printer_handler):
    try:
        hprinter = printer_handler
        printer_info = win32print.GetPrinter(hprinter, 2)
        printer_name = printer_info["pPrinterName"]
        print(f"Printing '{file_path}' on printer '{printer_name}'...")
        win32print.StartDocPrinter(hprinter, 1, ("Print Job", None, "RAW"))
        win32print.StartPagePrinter(hprinter)
        with open(file_path, "rb") as f:
            win32print.WritePrinter(hprinter, f.read())
        win32print.EndPagePrinter(hprinter)
        win32print.EndDocPrinter(hprinter)
        print("File printed successfully.")
    except Exception as e:
        print(f"Error printing '{file_path}': {e}")

def main():
    output_dir = Path.cwd() / "output"
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.Folders("mokdimprinting@outlook.com").Folders("Inbox")
    messages = inbox.Items
    for message in messages:
        subject = message.Subject
        attachments = message.Attachments

        if attachments.Count > 0:
            target_folder = output_dir / str(subject)
            target_folder.mkdir(parents=True, exist_ok=True)
            for attachment in attachments:
                attachment.SaveAsFile(target_folder / attachment.FileName)
                cur_file_path = target_folder / attachment.FileName
                printer_handler = get_printer()
                if printer_handler:
                    print_file(cur_file_path, printer_handler)
                else:
                    print("Printer not available. Skipping printing.")

if __name__ == "__main__":
    main()
