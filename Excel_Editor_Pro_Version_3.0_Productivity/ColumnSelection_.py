from PyQt5.QtWidgets import(QDialog, QVBoxLayout, QLabel, QScrollArea, QWidget, QHBoxLayout,
                            QPushButton, QCheckBox, QDialogButtonBox)


class ColumnSelectionDialog(QDialog):
    def __init__(self, columns, default_selected=None, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Select Columns for Report")
        self.setModal(True)
        self.resize(400, 500)
        
        layout = QVBoxLayout()
        layout.addWidget(QLabel("Select columns to include in the DOC report:"))
        
        self.scroll = QScrollArea()
        self.scroll_widget = QWidget()
        self.scroll_layout = QVBoxLayout()
        
        self.checkboxes = []
        default_selected = [s.lower() for s in (default_selected or [])]
        
        for col in columns:
            cb = QCheckBox(str(col))
            if str(col).lower() in default_selected:
                cb.setChecked(True)
            self.scroll_layout.addWidget(cb)
            self.checkboxes.append(cb)
            
        self.scroll_widget.setLayout(self.scroll_layout)
        self.scroll.setWidget(self.scroll_widget)
        self.scroll.setWidgetResizable(True)
        layout.addWidget(self.scroll)
        
        # Select All / Deselect All buttons
        btn_layout = QHBoxLayout()
        select_all_btn = QPushButton("Select All")
        select_all_btn.clicked.connect(self.select_all)
        deselect_all_btn = QPushButton("Deselect All")
        deselect_all_btn.clicked.connect(self.deselect_all)
        btn_layout.addWidget(select_all_btn)
        btn_layout.addWidget(deselect_all_btn)
        layout.addLayout(btn_layout)
        
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)
        
        self.setLayout(layout)
        
    def select_all(self):
        for cb in self.checkboxes:
            cb.setChecked(True)
            
    def deselect_all(self):
        for cb in self.checkboxes:
            cb.setChecked(False)
            
    def get_selected_columns(self):
        return [cb.text() for cb in self.checkboxes if cb.isChecked()]