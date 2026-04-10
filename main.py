import sys
import os
import win32com.client as win32
from enum import Enum
from emailscan import scan

class GUI(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Email Scanner")

        # Welcome text
        self.text = QtWidgets.QLabel("Welcome! Please select a date, or leave blank for all emails.", alignment=QtCore.Qt.AlignCenter)
        self.folder_name = QtWidgets.QTextEntry("Type subfolder (leave blank for main inbox).");

        # Button to scan
        self.scan_button = QtWidgets.QPushButton("Scan")
        self.scan_button.clicked.connect(emailscan.scan(str(self.folder_name)))

        # Layout for the main window
        self.layout = QtWidgets.QVBoxLayout(self)
        self.layout.addWidget(self.text)
        self.layout.addWidget(self.folder_name)
        self.layout.addWidget(self.scan_button)

if __name__ == "__main__":
    app = QtWidgets.QApplication([])
    widget = GUI()
    widget.resize(400, 250)
    widget.show()
    sys.exit(app.exec())

        file.write("-" * 100 + "\n\n")
        print(f"Processed unread email from {sender}")
