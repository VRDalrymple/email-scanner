import sys
import os
import json
import win32com.client as win32
from enum import Enum
import pandas as pd
import cv2
from PIL import Image
import pytesseract
import platform

def scan(image):
	if platform.system() == "Windows":
    pytesseract.pytesseract.tesseract_cmd = r'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'
	print(image)
	text = ''
	image = str(image)

	if image.endswith('.jpg') or image.endswith('.jpeg') or image.endswith('.png'):
			image = cv2.imread(image)
			try:
					image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
					_, threshold_image = cv2.threshold(image, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
					image = cv2.fastNlMeansDenoising(threshold_image, None, 30, 7, 21)
			except Exception as err:
					print(f'Error: {err}')
			try:
					custom_config = r'--oem 3 --psm 3'
					text_data = pytesseract.image_to_string(image, config=custom_config)
					text += text_data
			except Exception as err:
					print(f'Error: {err}')

	return text


class OutlookFolder(Enum):
    olFolderDeletedItems = 3 # The Deleted Items folder
    olFolderOutbox = 4 # The Outbox folder
    olFolderSentMail = 5 # The Sent Mail folder
    olFolderInbox = 6 # The Inbox folder
    olFolderDrafts = 16 # The Drafts folder
    olFolderJunk = 23 # The Junk E-Mail folder

with open("vardata.json", "r") as f:
	vardata = json.load(f)
  
# Get a reference to Outlook
outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Get the Inbox folder
inbox = outlook.GetDefaultFolder(OutlookFolder.olFolderInbox.value)

# Get subfolder
try:
	if vardata["folder"] != "":
		folder = inbox.Folders.Item(f"{vardata["folder"]}")
	else:
		messages = inbox.Items
except Exception as err:
  print(f"Error: {err}"}
	print("Subfolder not found. Scanning from main inbox.")
	messages = inbox.Items

# Apply restrictions
try:
	for restriction in vardata["restrictions"]:
		try:
		  if vardata["restrictions"][restriction] != "":
		    messages = folder.Items.Restrict(f"{vardata["restrictions"][restriction]}")
		except Exception as err:
		  print(f"Error: {err}"}
		  print("Restriction failed.")
except Exception as err:
	  print(f"Error: {err}"}
	  print("Restrictions failed.")

# Check if messages exist
if len(messages) == 0:
    print("There aren't any messages to scan in this folder")
    exit()

# Initialize text file
output_file = "OutreachResults.txt"
with open(output_file, "w", encoding="utf-8") as file:
    # Loop over messages
    for message in messages:
        sender = message.SenderName.upper()
        date_sent = message.Senton.date()
        email_body = message.Body
        attachments = message.Attachments
        
        file.write(f"Name: {sender}\nDate Sent: {date_sent}\n\nEmail Content:\n{email_body}\n")
        
        if len(attachments) > 0:
            for attachment in attachments:
                attachment_path = os.path.join(os.getcwd(), attachment.FileName)
                attachment.SaveASFile(attachment_path)
                scanned = scan(attachment_path)
                
                if scanned:
                    file.write("Attachment Content:\n")
                    file.write(scanned + "\n\n")  # Directly write the full text
        
        file.write("-_" * 100 + "-\n\n")
        print(f"Processed unread email from {sender}")
