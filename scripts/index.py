from pathlib import Path
import win32com.client
import win32print
import win32ui

# Printer name: HP LaserJet MFP M426fdn (5F550F)
from printerconfig import get_printer
printer = get_printer()

#create output folder
output_dir = Path.cwd() / "..//output"
# output_dir.mkdir(parents=True, exist_ok=True)

# Connect to Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Connect to folder
# inbox = outlook.GetDefaultFolder(6)

# RUN CODE EVERY 1 MINUTES TO CHECK FOR NEW EMAIL OR PING SERVER WHEN NEW EMAIL ENTERS SYSTME
# LOG INFORMATION ABOUT THE PROCESS
inbox = outlook.Folders("mokdimprinting@outlook.com").Folders("Inbox")

# Get messages
messages = inbox.Items

# Iterate over messages
for message in messages:
    subject = message.Subject
    body = message.body
    attachments = message.Attachments
    
    if (attachments.count > 0): # save message only if it has attachments
        # Save message and files information
        target_folder = output_dir / str(subject)
        target_folder.mkdir(parents=True, exist_ok=True)
        
        # Write body to text file (FUTURE DEVELOPMENT)
        # Path(target_folder / "body.txt").write_text(str(body))
        
        # Save attachments
        for attachment in attachments:
            attachment.SaveAsFile(target_folder / str(attachment))
            cur_file = open(Path.cwd() / "..//output" / target_folder / attachment.FileName, 'r')
            printer.StartDoc(cur_file)
            printer.StartPage()
            printer.EndPage()
            printer.EndDoc()
            
        # Print attachments
        
        # Delete attachments from output folder if print was successful
        
        # Delete file if print was successful