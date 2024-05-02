from pathlib import Path
import win32com.client

#create output folder
output_dir = Path.cwd() / "..//output"
# output_dir.mkdir(parents=True, exist_ok=True)

# Connect to Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Connect to folder
# inbox = outlook.GetDefaultFolder(6)
inbox = outlook.Folders("mokdimprinting@outlook.com").Folders("Inbox")

# Get messages
messages = inbox.Items

# Iterate over messages
for message in messages:
    subject = message.Subject
    body = message.body
    attachments = message.Attachments
    
    # Save message and files information
    target_folder = output_dir / str(subject)
    target_folder.mkdir(parents=True, exist_ok=True)
    
    # Write body to text file (FUTURE DEVELOPMENT)
    Path(target_folder / "EMAIL_BODY.txt").write_text(str(body))
    
    # Save attachments
    for attachment in attachments:
        attachment.SaveAsFile(target_folder / str(attachment))