import sys
from PySide6 import QtCore, QtWidgets
from PySide6.QtWidgets import (
    QApplication, QWidget, QLineEdit, QDateEdit,
    QLabel, QFormLayout, QMessageBox, QProgressBar
)
from PySide6.QtCore import QDate, QThread, QObject, Signal, Slot
from scanner import get_messages, email_scan

class Worker(QObject):
    finished = Signal()
    progress = Signal(int)
    results = Signal(object)
    error = Signal(str)

    def __init__(self, folder, date1, date2):
        super().__init__()
        self.folder = folder
        self.date1 = date1
        self.date2 = date2

    @Slot()
    def run(self):
        try:
            messages = get_messages(self.folder, self.date1, self.date2)

            if not messages:
                self.results.emit([])
                return

            email_scan(messages, self.progress.emit)
            self.results.emit(messages)

        except Exception as e:
            self.error.emit(str(e))

        finally:
            self.finished.emit()

class GUI(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Email Scanner")
        self.messages = []

        # Label
        self.text = QLabel(
            "Welcome!\n(Leave subfolder blank for main inbox)",
            alignment=QtCore.Qt.AlignCenter
        )

        # Subfolder input
        self.folder_name = QLineEdit("", parent=self)

        # Dates
        self.date1 = QDateEdit(self)
        self.date1.setDate(QDate.currentDate())

        self.date2 = QDateEdit(self)
        self.date2.setDate(QDate.currentDate())

        # Button
        self.scan_button = QtWidgets.QPushButton("Scan")
        self.scan_button.clicked.connect(self.start_scan)

        # Progress bar (0–100)
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)

        # Layout
        layout = QFormLayout()
        self.setLayout(layout)

        layout.addRow(self.text)
        layout.addRow('Subfolder:', self.folder_name)
        layout.addRow('Start Date:', self.date1)
        layout.addRow('End Date:', self.date2)
        layout.addRow(self.scan_button)
        layout.addRow(self.progress_bar)

    def start_scan(self):
        folder = self.folder_name.text().strip()

        # Clean date formatting (no manual parsing nonsense)
        date1 = self.date1.date().toString("yyyy-MM-dd")
        date2 = self.date2.date().toString("yyyy-MM-dd")

        # Create thread + worker
        self.thread = QThread()
        self.worker = Worker(folder, date1, date2)

        self.worker.moveToThread(self.thread)

        # Start
        self.thread.started.connect(self.worker.run)

        # Cleanup
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)

        # Signals → GUI
        self.worker.progress.connect(self.update_progress)
        self.worker.results.connect(self.handle_results)
        self.worker.error.connect(self.show_error)
        self.worker.finished.connect(self.done)

        # Reset progress
        self.progress_bar.setValue(0)

        self.thread.start()

    def update_progress(self, value):
        self.progress_bar.setValue(value)

    def handle_results(self, messages):
        if not messages:
            QMessageBox.information(self, "No Messages", "There are no messages in this folder.")
            return

        self.messages = messages

    def show_error(self, message):
        QMessageBox.critical(self, "Error", message)

    def done(self):
        QMessageBox.information(self, "Done!", "Done scanning!")

if __name__ == "__main__":
    app = QApplication([])
    widget = GUI()
    widget.show()
    sys.exit(app.exec())
