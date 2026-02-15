from PyQt5.QtWidgets import (QVBoxLayout, QLineEdit, QLabel, QComboBox, QDialog, 
                             QDialogButtonBox)

class AddColumnDialog(QDialog):
    def __init__(self, existing_columns, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Add New Column")
        self.setModal(True)
        self.resize(350, 200)
        
        layout = QVBoxLayout()
        
        layout.addWidget(QLabel("Column Name:"))
        self.column_name_edit = QLineEdit()
        self.column_name_edit.setPlaceholderText("Enter column name...")
        layout.addWidget(self.column_name_edit)
        
        layout.addWidget(QLabel("Insert Position:"))
        self.position_combo = QComboBox()
        positions = ["At End"] + [f"Before '{col}'" for col in existing_columns]
        self.position_combo.addItems(positions)
        layout.addWidget(self.position_combo)
        
        layout.addWidget(QLabel("Default Value (optional):"))
        self.default_value_edit = QLineEdit()
        self.default_value_edit.setPlaceholderText("Leave empty for blank cells...")
        layout.addWidget(self.default_value_edit)
        
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)
        
        self.setLayout(layout)
    
    def get_data(self):
        return {
            'name': self.column_name_edit.text(),
            'position': self.position_combo.currentIndex(),
            'default_value': self.default_value_edit.text()
        }