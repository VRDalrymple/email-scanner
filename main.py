import sys
from PySide6 import QtCore, QtWidgets
from PySide6.QtWidgets import QApplication, QWidget, QLineEdit, QDateEdit, QLabel, QFormLayout
from PySide6.QtCore import QDate, QDateTime, QThread, SIGNAL
from PySide6.QtWidgets import QMessageBox, QProgressBar
import datetime
from scanner import get_messages, email_scan

class Thread(QThread):
    def __init__(self):
        QThread.__init__(self)

    def __del__(self):
        self.wait()

    def run(self):
        messages = get_messages(widget.folder_name, widget.date1, widget.date2)

        if not messages:
            return

class GUI(QtWidgets.QWidget):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.setWindowTitle("Email Scanner")
        self.messages = 0

        # Thread
        self.thread = Thread()

        # Label
        self.text = QtWidgets.QLabel("Welcome!\n(Leave subfolder blank for main inbox)", alignment=QtCore.Qt.AlignCenter)
        
        # Subfolder
        self.folder_name = QLineEdit("", parent=self)
        self.folder_name.textEdited.connect(self.update_subfolder)
        
        # Date
        self.date1 = QtWidgets.QDateEdit(self)
        self.date1.setDate(QDate.currentDate())
        self.value1 = self.date1.date()
        self.value1 = self.format_date(self.value1)
        self.date1.editingFinished.connect(self.update_date1)

        self.date2 = QtWidgets.QDateEdit(self)
        self.date2.setDate(QDate.currentDate())
        self.value2 = self.date2.date()
        self.value2 = self.format_date(self.value2)
        self.date2.editingFinished.connect(self.update_date2)
        
        # Button to scan
        self.scan_button = QtWidgets.QPushButton("Scan")
        self.scan_button.clicked.connect(self.thread.start)

        # Progress bar
        self.progress_bar = QtWidgets.QProgressBar()
        self.connect(self.thread, SIGNAL("finished()"), self.done)
        
        # Layout
        self.layout = QFormLayout()
        self.setLayout(self.layout)

        self.layout.addRow(self.text)
        self.layout.addRow('Subfolder:', self.folder_name)
        self.layout.addRow('Start Date:', self.date1)
        self.layout.addRow('End Date:', self.date2)
        self.layout.addRow(self.scan_button)
        self.layout.addRow(self.progress_bar)

    def update_subfolder(self):
        self.folder_name.setText(self.folder_name.text())

    def format_date(self, value):
        # format - 'yyyy-mm-dd'
        new_value = value.toString().split()
        
        match new_value[1]:
            case "Jan":
                new_value[1] = "01"
            case "Feb":
                new_value[1] = "02"
            case "Mar":
                new_value[1] = "03"
            case "Apr":
                new_value[1] = "04"
            case "May":
                new_value[1] = "05"
            case "Jun":
                new_value[1] = "06"
            case "Jul":
                new_value[1] = "07"
            case "Aug":
                new_value[1] = "08"
            case "Sep":
                new_value[1] = "09"
            case "Oct":
                new_value[1] = "10"
            case "Nov":
                new_value[1] = "11"
            case "Dec":
                new_value[1] = "12"

        if int(new_value[2]) < 10:
            new_value[2] = f"0{new_value[2]}"

        value = [str(new_value[3]),str(new_value[1]),str(new_value[2])]
        value = "-".join(value)
        return value

    def update_date1(self):
            self.value1 = self.date1.date()
            self.value1 = self.format_date(self.value1)
            print(self.value1)
    
    def update_date2(self):
            self.value2 = self.date2.date()
            self.value2 = self.format_date(self.value2)
            print(self.value2)

    def make_progress(self):
        self.progress_bar.setMaximum(len(self.messages))

    def done(self):
        QMessageBox.information(self, "Done!", "Done scanning!")

    def handle_results(self, messages):
        self.messages = messages
        self.progress_bar.setMaximum(len(messages))
        self.progress_bar.setValue(0)

        email_scan(self.messages, self.progress_bar)
    
    def no_messages(self):
        QMessageBox.information(widget, "No Messages", "There are no messages in this folder.")
    
def run_scan():
    email_scan(widget.messages,widget.progress_bar)

if __name__ == "__main__":
    app = QtWidgets.QApplication([])
    widget = GUI()
    widget.show()
    sys.exit(app.exec())
