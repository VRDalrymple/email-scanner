import sys
import os
import re
import fitz
import tempfile
import win32com.client as win32
from enum import Enum
from ocr import image_scan

class OutlookFolder(Enum):
    olFolderDeletedItems = 3 # The Deleted Items folder
    olFolderOutbox = 4 # The Outbox folder
    olFolderSentMail = 5 # The Sent Mail folder
    olFolderInbox = 6 # The Inbox folder
    olFolderDrafts = 16 # The Drafts folder
    olFolderJunk = 23 # The Junk E-Mail folder

def uniquename(base_path):
    """Return a unique filename if base_path already exists."""
    counter = 1
    path, ext = os.path.splitext(base_path)
    candidate = base_path
    while os.path.exists(candidate):
        candidate = f"{path}_{counter}{ext}"
        counter += 1
    return candidate

def sanitize_filename(self, name, max_length=150):
    """Replace illegal filename characters and truncate safely."""
    safe = re.sub(r'[\\/*?:"<>|]', "_", name)
    return safe[:max_length]

def get_messages(sub,date1,date2):
    # Get a reference to Outlook
    outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    
    # Get the Inbox folder
    inbox = outlook.GetDefaultFolder(OutlookFolder.olFolderInbox.value)
    
    # For subfolders
    if sub != "":
        try:
            inbox = inbox.Folders.Item(f"{sub}")
        except BaseException as err:
            print("Invalid subfolder name. Defaulting to main inbox.")
            print({err})
            inbox = outlook.GetDefaultFolder(OutlookFolder.olFolderInbox.value)
    
    print(f"{sub}")
    if date1 == date2:
        restriction = f"[ReceivedTime]>='{date1}'"
    else:
        restriction = f"[ReceivedTime]>='{date1}' AND [ReceivedTime]<='{date2}'"

    messages = inbox.Items.Restrict(restriction)
    messages = [msg for msg in messages]

    return messages

def email_scan(messages, progress_callback=None):
    total = len(messages)
    processed = 0

    # Initialize text file
    output_file = "OutreachResults.txt"
    with open(output_file, "w", encoding="utf-8") as file:
        attachment_path = os.path.join(os.getcwd(), f"images")
        if os.path.exists(attachment_path) == False:
            os.mkdir(attachment_path)
        path = attachment_path
        # Loop over messages
        for message in messages:
            subject = message.Subject
            sender = message.SenderName.upper()
            date_sent = message.Senton.date()
            email_body = message.Body
            attachments = message.Attachments
            
            file.write(f"Subject: {subject}\nSender: {sender}\nDate Sent: {date_sent}\n\nEmail Content:\n{email_body}\n\n")
            
            if len(attachments) > 0:
                for attachment in attachments:
                    try:
                        filename = attachment.FileName.lower()
                        if filename.endswith(".pdf"):
                                attachment_path = os.path.join(path, attachment.FileName)
                                attachment.SaveASFile(attachment_path)

                                pdf = fitz.open(attachment_path)
                                for page_index in range(len(pdf)):
                                    page = pdf.load_page(page_index)
                                    image_list = page.get_images(full=True)

                                    if image_list:
                                        print(f"[+] Found {len(image_list)} images on page {page_index}")
                                    else:
                                        print(f"[!] No images found on page {page_index}")

                                    for image_index, img in enumerate(image_list, start=1):
                                        xref = img[0]
                                        base_image = pdf.extract_image(xref)
                                        image_bytes = base_image["image"]
                                        image_ext = base_image["ext"]

                                        image_name = os.path.join(
                                            path,
                                            f"{subject}_{page_index+1}_{image_index}.{image_ext}"
                                        )
                                        image_name = uniquename(image_name)
                                        with open(image_name, "wb") as image_file:
                                            image_file.write(image_bytes)

                                        print(f"[+] Extracted image saved as {image_name}")
                                        attachment_scanned = image_scan(image_name)

                                        saved_any = True

                                pdf.close()
                        else:
                            attachment_path = os.path.join(path, attachment.FileName)
                            attachment_path = uniquename(attachment.FileName)
                            attachment.SaveASFile(attachment_path)
                            attachment_scanned = image_scan(attachment_path)
                            
                        if attachment_scanned:
                            file.write("Attachment Content:\n")
                            file.write(attachment_scanned + "\n\n")  # Directly write the full text
                    except:
                        print("couldn't process attachment")
        
            file.write("-" * 100 + "\n\n")
            print(f"Processed unread email from {sender}")
            processed += 1
            if progress_callback:
                percent = int((processed / total) * 100)
                progress_callback(percent)
