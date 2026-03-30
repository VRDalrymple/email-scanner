import sys
import os
import win32com.client as win32
from enum import Enum
from ocr import scan

class OutlookFolder(Enum):
    olFolderDeletedItems = 3 # The Deleted Items folder
    olFolderOutbox = 4 # The Outbox folder
    olFolderSentMail = 5 # The Sent Mail folder
    olFolderInbox = 6 # The Inbox folder
    olFolderDrafts = 16 # The Drafts folder
    olFolderJunk = 23 # The Junk E-Mail folder

# Get a reference to Outlook
outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Get the Inbox folder
inbox = outlook.GetDefaultFolder(OutlookFolder.olFolderInbox.value)

# Get subfolder "Reps"
reps = inbox.Folders.Item("Reps")

# Apply the restriction
messages = reps.Items.Restrict("[ReceivedTime]>='2025-05-01'")

# Check if messages exist
if len(messages) == 0:
    print("There aren't any unread messages in this folder")
    exit()

# Initialize text file
output_file = "OutreachResults.txt"
with open(output_file, "w", encoding="utf-8") as file:
    # Loop over messages
    for message in messages:
        name = message.SenderName.upper()
        date_sent = message.Senton.date()
        email_body = message.Body
        attachments = message.Attachments
        print(name)
        
        file.write(f"Name: {name}\nDate Sent: {date_sent}\n\nEmail Content:\n{email_body}\n\n")
        
        if len(attachments) > 0:
            for attachment in attachments:
                attachment_path = os.path.join(os.getcwd(), attachment.FileName)
                attachment.SaveASFile(attachment_path)
                scanned = scan(attachment_path)
                
                if scanned:
                    file.write("Attachment Content:\n")
                    file.write(scanned + "\n\n")  # Directly write the full text


        
        file.write("-" * 100 + "\n\n")
        print(f"Processed unread email from {repname}")
