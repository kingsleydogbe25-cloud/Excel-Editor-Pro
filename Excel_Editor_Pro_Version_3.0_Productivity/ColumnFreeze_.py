"""
Column/Row Freezing Module
Allows freezing of header rows and columns for easier navigation
"""

from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QPushButton,
                             QLabel, QSpinBox, QGroupBox, QCheckBox, QMessageBox)
from PyQt5.QtCore import QSettings


class ColumnFreezeManager:
    """Manages column and row freezing"""
    
    def __init__(self, parent):
        self.parent = parent
        self.settings = QSettings("DataTools", "ExcelEditor")
        
        # Freeze settings
        self.freeze_rows = 0
        self.freeze_cols = 0
        self.freeze_enabled = False
        
        # Load saved settings
        self.load_settings()
    
    def load_settings(self):
        """Load freeze settings"""
        self.freeze_rows = int(self.settings.value("freeze/rows", 0))
        self.freeze_cols = int(self.settings.value("freeze/cols", 0))
        self.freeze_enabled = self.settings.value("freeze/enabled", False, type=bool)
    
    def save_settings(self):
        """Save freeze settings"""
        self.settings.setValue("freeze/rows", self.freeze_rows)
        self.settings.setValue("freeze/cols", self.freeze_cols)
        self.settings.setValue("freeze/enabled", self.freeze_enabled)
    
    def apply_freeze(self, rows=None, cols=None):
        """Apply freezing to table"""
        try:
            if rows is not None:
                self.freeze_rows = rows
            if cols is not None:
                self.freeze_cols = cols
            
            table = self.parent.table_widget
            
            if self.freeze_enabled and (self.freeze_rows > 0 or self.freeze_cols > 0):
                # Freeze rows (make them sticky at top)
                if self.freeze_rows > 0:
                    # Set vertical header to show frozen rows
                    for row in range(min(self.freeze_rows, table.rowCount())):
                        # Make row header bold
                        header_item = table.verticalHeaderItem(row)
                        if header_item:
                            from PyQt5.QtGui import QFont
                            font = header_item.font()
                            font.setBold(True)
                            header_item.setFont(font)
                
                # Freeze columns (make them sticky at left)
                if self.freeze_cols > 0:
                    # Make column headers bold
                    for col in range(min(self.freeze_cols, table.columnCount())):
                        header_item = table.horizontalHeaderItem(col)
                        if header_item:
                            from PyQt5.QtGui import QFont
                            font = header_item.font()
                            font.setBold(True)
                            header_item.setFont(font)
                
                # Update table styling to show frozen sections
                self.update_frozen_styling()
                
                self.parent.statusBar().showMessage(
                    f"Frozen {self.freeze_rows} row(s) and {self.freeze_cols} column(s)",
                    3000
                )
            else:
                # Remove freeze
                self.remove_freeze_styling()
                self.parent.statusBar().showMessage("Freeze removed", 2000)
            
            self.save_settings()
            
        except Exception as e:
            print(f"Error applying freeze: {e}")
            QMessageBox.critical(
                self.parent,
                "Freeze Error",
                f"Failed to apply freeze:\n{str(e)}"
            )
    
    def update_frozen_styling(self):
        """Update styling for frozen rows/columns"""
        try:
            table = self.parent.table_widget
            
            # Style frozen rows
            for row in range(min(self.freeze_rows, table.rowCount())):
                for col in range(table.columnCount()):
                    item = table.item(row, col)
                    if item:
                        from PyQt5.QtGui import QColor, QBrush, QFont
                        # Light blue background for frozen cells
                        item.setBackground(QBrush(QColor(230, 240, 255)))
                        font = item.font()
                        font.setBold(True)
                        item.setFont(font)
            
            # Style frozen columns
            for col in range(min(self.freeze_cols, table.columnCount())):
                for row in range(table.rowCount()):
                    # Skip cells already styled as frozen rows
                    if row >= self.freeze_rows:
                        item = table.item(row, col)
                        if item:
                            from PyQt5.QtGui import QColor, QBrush, QFont
                            # Light yellow background for frozen columns
                            item.setBackground(QBrush(QColor(255, 250, 230)))
                            font = item.font()
                            font.setBold(True)
                            item.setFont(font)
            
        except Exception as e:
            print(f"Error updating frozen styling: {e}")
    
    def remove_freeze_styling(self):
        """Remove freeze styling from table"""
        try:
            table = self.parent.table_widget
            
            # Reset all cell backgrounds and fonts
            for row in range(table.rowCount()):
                for col in range(table.columnCount()):
                    item = table.item(row, col)
                    if item:
                        from PyQt5.QtGui import QBrush, QFont
                        # Reset to default
                        item.setBackground(QBrush())
                        font = item.font()
                        font.setBold(False)
                        item.setFont(font)
            
            # Reset headers
            for row in range(table.rowCount()):
                header_item = table.verticalHeaderItem(row)
                if header_item:
                    font = header_item.font()
                    font.setBold(False)
                    header_item.setFont(font)
            
            for col in range(table.columnCount()):
                header_item = table.horizontalHeaderItem(col)
                if header_item:
                    font = header_item.font()
                    font.setBold(False)
                    header_item.setFont(font)
            
        except Exception as e:
            print(f"Error removing freeze styling: {e}")
    
    def toggle_freeze(self):
        """Toggle freeze on/off"""
        self.freeze_enabled = not self.freeze_enabled
        self.apply_freeze()
    
    def freeze_top_row(self):
        """Quick freeze: freeze top row only"""
        self.freeze_enabled = True
        self.apply_freeze(rows=1, cols=0)
    
    def freeze_first_column(self):
        """Quick freeze: freeze first column only"""
        self.freeze_enabled = True
        self.apply_freeze(rows=0, cols=1)
    
    def freeze_both(self):
        """Quick freeze: freeze both top row and first column"""
        self.freeze_enabled = True
        self.apply_freeze(rows=1, cols=1)
    
    def unfreeze_all(self):
        """Remove all freezing"""
        self.freeze_enabled = False
        self.freeze_rows = 0
        self.freeze_cols = 0
        self.apply_freeze()


class ColumnFreezeDialog(QDialog):
    """Dialog for configuring column/row freezing"""
    
    def __init__(self, parent, freeze_manager):
        super().__init__(parent)
        self.parent = parent
        self.freeze_manager = freeze_manager
        self.init_ui()
    
    def init_ui(self):
        self.setWindowTitle("Freeze Panes")
        self.setGeometry(300, 300, 400, 350)
        
        layout = QVBoxLayout()
        
        # Description
        desc_label = QLabel(
            "Freeze rows and columns to keep them visible while scrolling.\n"
            "Frozen sections will be highlighted."
        )
        desc_label.setWordWrap(True)
        layout.addWidget(desc_label)
        
        # Enable checkbox
        self.enable_checkbox = QCheckBox("Enable Freeze Panes")
        self.enable_checkbox.setChecked(self.freeze_manager.freeze_enabled)
        self.enable_checkbox.stateChanged.connect(self.on_enable_changed)
        layout.addWidget(self.enable_checkbox)
        
        # Freeze settings group
        freeze_group = QGroupBox("Freeze Settings")
        freeze_layout = QVBoxLayout()
        
        # Rows
        rows_layout = QHBoxLayout()
        rows_layout.addWidget(QLabel("Freeze top rows:"))
        
        self.rows_spinbox = QSpinBox()
        self.rows_spinbox.setMinimum(0)
        self.rows_spinbox.setMaximum(100)
        self.rows_spinbox.setValue(self.freeze_manager.freeze_rows)
        rows_layout.addWidget(self.rows_spinbox)
        rows_layout.addStretch()
        
        freeze_layout.addLayout(rows_layout)
        
        # Columns
        cols_layout = QHBoxLayout()
        cols_layout.addWidget(QLabel("Freeze left columns:"))
        
        self.cols_spinbox = QSpinBox()
        self.cols_spinbox.setMinimum(0)
        self.cols_spinbox.setMaximum(50)
        self.cols_spinbox.setValue(self.freeze_manager.freeze_cols)
        cols_layout.addWidget(self.cols_spinbox)
        cols_layout.addStretch()
        
        freeze_layout.addLayout(cols_layout)
        
        freeze_group.setLayout(freeze_layout)
        layout.addWidget(freeze_group)
        
        # Quick actions group
        quick_group = QGroupBox("Quick Actions")
        quick_layout = QVBoxLayout()
        
        freeze_header_btn = QPushButton("Freeze Header Row (Top Row)")
        freeze_header_btn.clicked.connect(self.freeze_header)
        quick_layout.addWidget(freeze_header_btn)
        
        freeze_first_col_btn = QPushButton("Freeze First Column")
        freeze_first_col_btn.clicked.connect(self.freeze_first_column)
        quick_layout.addWidget(freeze_first_col_btn)
        
        freeze_both_btn = QPushButton("Freeze Both (Top Row + First Column)")
        freeze_both_btn.clicked.connect(self.freeze_both)
        quick_layout.addWidget(freeze_both_btn)
        
        unfreeze_btn = QPushButton("Unfreeze All")
        unfreeze_btn.clicked.connect(self.unfreeze_all)
        quick_layout.addWidget(unfreeze_btn)
        
        quick_group.setLayout(quick_layout)
        layout.addWidget(quick_group)
        
        layout.addStretch()
        
        # Buttons
        button_layout = QHBoxLayout()
        
        apply_btn = QPushButton("Apply")
        apply_btn.clicked.connect(self.apply_settings)
        button_layout.addWidget(apply_btn)
        
        ok_btn = QPushButton("OK")
        ok_btn.clicked.connect(self.accept_settings)
        button_layout.addWidget(ok_btn)
        
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        button_layout.addWidget(cancel_btn)
        
        layout.addLayout(button_layout)
        
        self.setLayout(layout)
    
    def on_enable_changed(self, state):
        """Handle enable checkbox change"""
        enabled = state == 2  # Qt.Checked
        self.rows_spinbox.setEnabled(enabled)
        self.cols_spinbox.setEnabled(enabled)
    
    def apply_settings(self):
        """Apply current settings"""
        self.freeze_manager.freeze_enabled = self.enable_checkbox.isChecked()
        self.freeze_manager.apply_freeze(
            rows=self.rows_spinbox.value(),
            cols=self.cols_spinbox.value()
        )
    
    def accept_settings(self):
        """Apply and close"""
        self.apply_settings()
        self.accept()
    
    def freeze_header(self):
        """Quick action: freeze header row"""
        self.enable_checkbox.setChecked(True)
        self.rows_spinbox.setValue(1)
        self.cols_spinbox.setValue(0)
        self.apply_settings()
    
    def freeze_first_column(self):
        """Quick action: freeze first column"""
        self.enable_checkbox.setChecked(True)
        self.rows_spinbox.setValue(0)
        self.cols_spinbox.setValue(1)
        self.apply_settings()
    
    def freeze_both(self):
        """Quick action: freeze both"""
        self.enable_checkbox.setChecked(True)
        self.rows_spinbox.setValue(1)
        self.cols_spinbox.setValue(1)
        self.apply_settings()
    
    def unfreeze_all(self):
        """Quick action: unfreeze all"""
        self.enable_checkbox.setChecked(False)
        self.rows_spinbox.setValue(0)
        self.cols_spinbox.setValue(0)
        self.apply_settings()
