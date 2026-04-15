# Email Scanner (Outlook + OCR)

A PySide6 desktop application that scans Outlook emails within a date range, extracts content (including attachments), and writes results to a text file.

---

## Features

* Scan Outlook Inbox or specific subfolders
* Filter emails by date range
* Extract:

  * Sender
  * Date sent
  * Email body
* Process attachments:

  * PDF image extraction
  * OCR on images
* Progress bar with live updates
* Output saved to `OutreachResults.txt`

---

## Tech Stack

* Python 3
* PySide6 (PyQt GUI)
* pywin32 (Outlook COM integration)
* PyMuPDF (`fitz`) for PDF processing
* PyTesseract

---

## Project Structure

```
.
├── main.py          # GUI + threading logic
├── scanner.py       # Outlook + scanning logic
├── ocr.py           # Image OCR processing
├── images/          # Extracted images (auto-created)
├── OutreachResults.txt
```

---

## Installation

### 1. Install dependencies

```bash
pip install PySide6 pywin32 pymupdf pytesseract
```

---

### 2. Windows Requirement

This app uses Outlook COM, so:

* Windows only
* Microsoft Outlook must be installed and configured

---

## Usage

Run:

```bash
python main.py
```

### In the app:

1. Enter subfolder (optional)

   * Leave blank for Inbox
   * Use exact Outlook folder name
   * Nested folders use `/`

     ```
     Projects/2026/Invoices
     ```

2. Select:

   * Start date
   * End date

3. Click **Scan**

---

## 📄 Output

Results are written to:

```
OutreachResults.txt
```

Includes:

* Sender
* Subject
* Date sent
* Email body
* Extracted attachment text (if any)

---

## How It Works

1. GUI collects user input
2. A background worker thread:

   * Fetches messages from Outlook
   * Filters by date
   * Passes messages to scanner
3. `email_scan()`:

   * Iterates messages
   * Saves attachments
   * Extracts images from PDFs
   * Runs OCR
4. Progress updates are emitted via Qt signals
