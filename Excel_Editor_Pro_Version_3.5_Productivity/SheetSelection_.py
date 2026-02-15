from PyQt5.QtWidgets import (QVBoxLayout,QLabel, QComboBox, QDialog, 
                             QDialogButtonBox)


class SheetSelectionDialog(QDialog):
    def __init__(self, sheet_names, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Select Sheet")
        self.setModal(True)
        self.resize(300, 200)
        
        layout = QVBoxLayout()
        layout.addWidget(QLabel("Multiple sheets found. Please select one:"))
        
        self.sheet_combo = QComboBox()
        self.sheet_combo.addItems(sheet_names)
        layout.addWidget(self.sheet_combo)
        
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)
        
        self.setLayout(layout)
    
    def get_selected_sheet(self):
        return self.sheet_combo.currentText()