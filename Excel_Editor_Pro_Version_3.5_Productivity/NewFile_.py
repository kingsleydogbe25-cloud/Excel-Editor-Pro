from PyQt5.QtWidgets import(QDialog, QVBoxLayout, QLabel, QComboBox, QGroupBox, QFormLayout,
                            QSpinBox, QCheckBox, QDialogButtonBox)


class NewFileDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Create New File")
        self.setModal(True)
        self.resize(350, 250)
        
        layout = QVBoxLayout()
        
        layout.addWidget(QLabel("Select file type:"))
        self.file_type_combo = QComboBox()
        self.file_type_combo.addItems(["Excel (.xlsx)", "CSV (.csv)"])
        layout.addWidget(self.file_type_combo)
        
        dimensions_group = QGroupBox("Initial Dimensions")
        dimensions_layout = QFormLayout()
        
        self.rows_spinbox = QSpinBox()
        self.rows_spinbox.setMinimum(1)
        self.rows_spinbox.setMaximum(10000)
        self.rows_spinbox.setValue(10)
        dimensions_layout.addRow("Rows:", self.rows_spinbox)
        
        self.cols_spinbox = QSpinBox()
        self.cols_spinbox.setMinimum(1)
        self.cols_spinbox.setMaximum(1000)
        self.cols_spinbox.setValue(5)
        dimensions_layout.addRow("Columns:", self.cols_spinbox)
        
        dimensions_group.setLayout(dimensions_layout)
        layout.addWidget(dimensions_group)
        
        self.include_headers_cb = QCheckBox("Include column headers")
        self.include_headers_cb.setChecked(True)
        layout.addWidget(self.include_headers_cb)
        
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)
        
        self.setLayout(layout)

    def get_settings(self):
        """Return the dialog settings"""
        return {
            'file_type': self.file_type_combo.currentText(),
            'rows': self.rows_spinbox.value(),
            'columns': self.cols_spinbox.value(),
            'include_headers': self.include_headers_cb.isChecked()
        }
