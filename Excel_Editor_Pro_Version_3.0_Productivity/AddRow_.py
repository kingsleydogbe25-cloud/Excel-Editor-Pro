from PyQt5.QtWidgets import(QDialog, QVBoxLayout, QScrollArea,QWidget, QFormLayout,
                            QLineEdit, QLabel, QDialogButtonBox)


class AddRowDialog(QDialog):
    def __init__(self, columns, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Add New Row")
        self.setModal(True)
        self.resize(400, min(600, 100 + len(columns) * 35))
        
        layout = QVBoxLayout()
        
        scroll = QScrollArea()
        scroll_widget = QWidget()
        
        form_layout = QFormLayout()
        self.inputs = {}
        
        for col in columns:
            line_edit = QLineEdit()
            form_layout.addRow(QLabel(str(col) + ":"), line_edit)
            self.inputs[col] = line_edit
        
        scroll_widget.setLayout(form_layout)
        scroll.setWidget(scroll_widget)
        scroll.setWidgetResizable(True)
        layout.addWidget(scroll)
        
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)
        
        self.setLayout(layout)
    def get_data(self):
        """Return dictionary of column names and their input values"""
        return {col: line_edit.text() for col, line_edit in self.inputs.items()}
