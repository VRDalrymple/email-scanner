from PySide6 import QtCore, QtWidgets
from scanner import scan

class GUI(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Email Scanner")

        self.text = QtWidgets.QLabel("Welcome!", alignment=QtCore.Qt.AlignCenter)
        
        # Subfolder
        self.folder_name = QLineEdit("Leave blank for main inbox", parent=self)
        self.folder_name.textEdited.connect(self.update_subfolder)
        
        # Date
        self.date = QtWidgets.QDateEdit(self)

        # Button to scan
        self.scan_button = QtWidgets.QPushButton("Scan")
        self.scan_button.clicked.connect(scan(self.folder_name.text(),str(self.date.date.toPyDate())))
        
        # Layout
        layout = QFormLayout()
        self.setLayout(layout)
        self.layout.addRow(self.text)
        self.layout.addRow('Subfolder:', self.folder_name)
        self.layout.addRow('Date:', self.date)
        self.layout.addRow(self.scan_button)

    def update_subfolder(self):
        self.label.setText(self.folder_name.text())

if __name__ == "__main__":
    app = QtWidgets.QApplication([])
    widget = GUI()
    widget.resize(400, 250)
    widget.show()
    sys.exit(app.exec())
