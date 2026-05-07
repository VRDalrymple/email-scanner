import os
import re
from enum import Enum

import fitz
import win32com.client as win32

from ocr import image_scan


class OutlookFolder(Enum):
    olFolderDeletedItems = 3
    olFolderOutbox = 4
    olFolderSentMail = 5
    olFolderInbox = 6
    olFolderDrafts = 16
    olFolderJunk = 23


def uniquename(base_path):
    """
    Return a unique filename if the path already exists.
    """

    counter = 1
    path, ext = os.path.splitext(base_path)

    candidate = base_path

    while os.path.exists(candidate):
        candidate = f"{path}_{counter}{ext}"
        counter += 1

    return candidate


def sanitize_filename(name, max_length=150):
    """
    Replace illegal Windows filename characters.
    """

    safe = re.sub(r'[\\/*?:"<>|]', "_", name)

    return safe[:max_length]


def get_messages(subfolder, date1, date2):

    outlook = win32.Dispatch(
        "Outlook.Application"
    ).GetNamespace("MAPI")

    inbox = outlook.GetDefaultFolder(
        OutlookFolder.olFolderInbox.value
    )

    # optional subfolder
    if subfolder:
        try:
            inbox = inbox.Folders.Item(subfolder)

        except Exception as err:
            print(
                f"Invalid subfolder '{subfolder}'. "
                f"Using main inbox instead."
            )
            print(err)

    # build restriction
    if date1 == date2:
        restriction = f"[ReceivedTime] >= '{date1}'"

    else:
        restriction = (
            f"[ReceivedTime] >= '{date1}' "
            f"AND [ReceivedTime] <= '{date2}'"
        )

    try:
        messages = inbox.Items.Restrict(restriction)

        # filter valid mail items only
        messages = [
            msg for msg in messages
            if hasattr(msg, "Subject")
        ]

        return messages

    except Exception as err:
        print(f"Failed retrieving messages: {err}")

        return []


def save_attachment(attachment, save_dir):

    safe_name = sanitize_filename(
        attachment.FileName
    )

    attachment_path = uniquename(
        os.path.join(save_dir, safe_name)
    )

    attachment.SaveASFile(attachment_path)

    print(f"[+] Saved attachment: {attachment_path}")

    return attachment_path


def extract_pdf_images(pdf_path, output_dir, subject):

    extracted_text = ""

    safe_subject = sanitize_filename(subject)

    try:
        pdf = fitz.open(pdf_path)

        for page_index in range(len(pdf)):

            page = pdf.load_page(page_index)

            image_list = page.get_images(full=True)

            if image_list:
                print(
                    f"[+] Found {len(image_list)} "
                    f"images on page {page_index + 1}"
                )
            else:
                print(
                    f"[!] No images found on "
                    f"page {page_index + 1}"
                )

            for image_index, img in enumerate(
                image_list,
                start=1
            ):

                try:
                    xref = img[0]

                    base_image = pdf.extract_image(xref)

                    image_bytes = base_image["image"]

                    # force png extension for OCR compatibility
                    image_name = uniquename(
                        os.path.join(
                            output_dir,
                            (
                                f"{safe_subject}_"
                                f"{page_index + 1}_"
                                f"{image_index}.png"
                            )
                        )
                    )

                    with open(image_name, "wb") as image_file:
                        image_file.write(image_bytes)

                    print(
                        f"[+] Extracted image saved: "
                        f"{image_name}"
                    )

                    scanned_text = image_scan(image_name)

                    if scanned_text:
                        extracted_text += (
                            scanned_text + "\n\n"
                        )

                except Exception as err:
                    print(
                        f"Failed extracting PDF image: {err}"
                    )

        pdf.close()

    except Exception as err:
        print(f"Failed processing PDF: {err}")

    return extracted_text


def email_scan(messages, progress_callback=None):

    total = len(messages)

    processed = 0

    output_file = "OutreachResults.txt"

    image_dir = os.path.join(
        os.getcwd(),
        "images"
    )

    os.makedirs(image_dir, exist_ok=True)

    with open(
        output_file,
        "w",
        encoding="utf-8"
    ) as file:

        for message in messages:

            try:
                subject = getattr(
                    message,
                    "Subject",
                    "No Subject"
                )

                sender = getattr(
                    message,
                    "SenderName",
                    "Unknown Sender"
                ).upper()

                email_body = getattr(
                    message,
                    "Body",
                    ""
                )

                date_sent = getattr(
                    message,
                    "SentOn",
                    None
                )

                if date_sent:
                    date_sent = date_sent.date()

                attachments = getattr(
                    message,
                    "Attachments",
                    []
                )

                file.write(
                    f"Subject: {subject}\n"
                    f"Sender: {sender}\n"
                    f"Date Sent: {date_sent}\n\n"
                    f"Email Content:\n"
                    f"{email_body}\n\n"
                )

                # process attachments
                if len(attachments) > 0:

                    for attachment in attachments:

                        try:
                            filename = (
                                attachment.FileName.lower()
                            )

                            attachment_text = ""

                            # PDF handling
                            if filename.endswith(".pdf"):

                                pdf_path = save_attachment(
                                    attachment,
                                    image_dir
                                )

                                attachment_text = (
                                    extract_pdf_images(
                                        pdf_path,
                                        image_dir,
                                        subject
                                    )
                                )

                            # image handling
                            elif filename.endswith(
                                (
                                    ".jpg",
                                    ".jpeg",
                                    ".png",
                                    ".bmp",
                                    ".tiff"
                                )
                            ):

                                image_path = save_attachment(
                                    attachment,
                                    image_dir
                                )

                                attachment_text = (
                                    image_scan(image_path)
                                )

                            else:
                                print(
                                    f"[!] Skipping unsupported "
                                    f"file type: {filename}"
                                )

                            if attachment_text:

                                file.write(
                                    "Attachment Content:\n"
                                )

                                file.write(
                                    attachment_text + "\n"
                                )

                        except Exception as err:
                            print(
                                f"Couldn't process "
                                f"attachment: {err}"
                            )

                file.write(
                    "-" * 100 + "\n\n"
                )

                print(
                    f"Processed unread email "
                    f"from {sender}"
                )

                processed += 1

                if progress_callback:

                    percent = int(
                        (processed / total) * 100
                    )

                    progress_callback(percent)

            except Exception as err:
                print(
                    f"Failed processing message: {err}"
                )
