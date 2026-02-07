from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, QComboBox, 
                             QSpinBox, QCheckBox, QPushButton, QGroupBox, QColorDialog,
                             QDialogButtonBox, QTabWidget, QWidget, QLineEdit)
from PyQt5.QtGui import QColor
from PyQt5.QtCore import Qt


class ColumnFormattingDialog(QDialog):
    """
    Advanced column formatting dialog for Excel files.
    Supports width, alignment, number formats, fonts, colors, and more.
    """
    def __init__(self, columns, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Column Formatting")
        self.setModal(True)
        self.resize(550, 600)
        
        self.columns = columns
        self.format_settings = {}
        
        # Initialize default colors
        self.bg_color = QColor(255, 255, 255)  # White
        self.text_color = QColor(0, 0, 0)  # Black
        self.header_bg_color = QColor(68, 114, 196)  # Blue
        self.header_text_color = QColor(255, 255, 255)  # White
        
        self.init_ui()
    
    def init_ui(self):
        main_layout = QVBoxLayout()
        
        # Column Selection
        col_group = QGroupBox("Select Column")
        col_layout = QHBoxLayout()
        col_layout.addWidget(QLabel("Column:"))
        self.column_combo = QComboBox()
        self.column_combo.addItems([str(col) for col in self.columns])
        self.column_combo.currentIndexChanged.connect(self.on_column_changed)
        col_layout.addWidget(self.column_combo)
        col_group.setLayout(col_layout)
        main_layout.addWidget(col_group)
        
        # Tab widget for different formatting options
        self.tabs = QTabWidget()
        
        # Tab 1: Basic Formatting
        basic_tab = QWidget()
        basic_layout = QVBoxLayout()
        
        # Width Settings
        width_group = QGroupBox("Column Width")
        width_layout = QHBoxLayout()
        self.auto_width_cb = QCheckBox("Auto-fit width")
        self.auto_width_cb.toggled.connect(self.toggle_width_input)
        width_layout.addWidget(self.auto_width_cb)
        width_layout.addWidget(QLabel("Custom Width:"))
        self.width_spin = QSpinBox()
        self.width_spin.setRange(5, 500)
        self.width_spin.setValue(100)
        self.width_spin.setSuffix(" px")
        width_layout.addWidget(self.width_spin)
        width_group.setLayout(width_layout)
        basic_layout.addWidget(width_group)
        
        # Alignment Settings
        align_group = QGroupBox("Alignment")
        align_layout = QVBoxLayout()
        
        h_align_layout = QHBoxLayout()
        h_align_layout.addWidget(QLabel("Horizontal:"))
        self.h_align_combo = QComboBox()
        self.h_align_combo.addItems(["Left", "Center", "Right", "General"])
        h_align_layout.addWidget(self.h_align_combo)
        align_layout.addLayout(h_align_layout)
        
        v_align_layout = QHBoxLayout()
        v_align_layout.addWidget(QLabel("Vertical:"))
        self.v_align_combo = QComboBox()
        self.v_align_combo.addItems(["Top", "Center", "Bottom"])
        self.v_align_combo.setCurrentIndex(1)  # Default to Center
        v_align_layout.addWidget(self.v_align_combo)
        align_layout.addLayout(v_align_layout)
        
        self.wrap_text_cb = QCheckBox("Wrap text")
        align_layout.addWidget(self.wrap_text_cb)
        
        align_group.setLayout(align_layout)
        basic_layout.addWidget(align_group)
        
        basic_layout.addStretch()
        basic_tab.setLayout(basic_layout)
        self.tabs.addTab(basic_tab, "Basic")
        
        # Tab 2: Number Formatting
        number_tab = QWidget()
        number_layout = QVBoxLayout()
        
        format_group = QGroupBox("Number Format")
        format_layout = QVBoxLayout()
        
        format_type_layout = QHBoxLayout()
        format_type_layout.addWidget(QLabel("Format Type:"))
        self.format_type_combo = QComboBox()
        self.format_type_combo.addItems([
            "General",
            "Number",
            "Currency",
            "Accounting",
            "Percentage",
            "Date",
            "Time",
            "Scientific",
            "Text"
        ])
        self.format_type_combo.currentIndexChanged.connect(self.on_format_type_changed)
        format_type_layout.addWidget(self.format_type_combo)
        format_layout.addLayout(format_type_layout)
        
        # Decimal places
        decimal_layout = QHBoxLayout()
        decimal_layout.addWidget(QLabel("Decimal Places:"))
        self.decimal_spin = QSpinBox()
        self.decimal_spin.setRange(0, 10)
        self.decimal_spin.setValue(2)
        decimal_layout.addWidget(self.decimal_spin)
        format_layout.addLayout(decimal_layout)
        
        # Currency symbol
        currency_layout = QHBoxLayout()
        currency_layout.addWidget(QLabel("Currency Symbol:"))
        self.currency_edit = QLineEdit()
        self.currency_edit.setText("$")
        self.currency_edit.setMaxLength(3)
        self.currency_edit.setEnabled(False)
        currency_layout.addWidget(self.currency_edit)
        format_layout.addLayout(currency_layout)
        
        # Use thousands separator
        self.thousands_cb = QCheckBox("Use thousands separator (,)")
        self.thousands_cb.setEnabled(False)
        format_layout.addWidget(self.thousands_cb)
        
        format_group.setLayout(format_layout)
        number_layout.addWidget(format_group)
        
        number_layout.addStretch()
        number_tab.setLayout(number_layout)
        self.tabs.addTab(number_tab, "Number Format")
        
        # Tab 3: Font and Colors
        style_tab = QWidget()
        style_layout = QVBoxLayout()
        
        # Font Settings
        font_group = QGroupBox("Font")
        font_layout = QVBoxLayout()
        
        font_name_layout = QHBoxLayout()
        font_name_layout.addWidget(QLabel("Font:"))
        self.font_combo = QComboBox()
        self.font_combo.addItems([
            "Calibri", "Arial", "Times New Roman", "Courier New", 
            "Verdana", "Tahoma", "Georgia", "Comic Sans MS"
        ])
        font_name_layout.addWidget(self.font_combo)
        font_layout.addLayout(font_name_layout)
        
        font_size_layout = QHBoxLayout()
        font_size_layout.addWidget(QLabel("Size:"))
        self.font_size_spin = QSpinBox()
        self.font_size_spin.setRange(6, 72)
        self.font_size_spin.setValue(11)
        font_size_layout.addWidget(self.font_size_spin)
        font_layout.addLayout(font_size_layout)
        
        font_style_layout = QHBoxLayout()
        self.bold_cb = QCheckBox("Bold")
        self.italic_cb = QCheckBox("Italic")
        self.underline_cb = QCheckBox("Underline")
        font_style_layout.addWidget(self.bold_cb)
        font_style_layout.addWidget(self.italic_cb)
        font_style_layout.addWidget(self.underline_cb)
        font_layout.addLayout(font_style_layout)
        
        font_group.setLayout(font_layout)
        style_layout.addWidget(font_group)
        
        # Color Settings
        color_group = QGroupBox("Colors")
        color_layout = QVBoxLayout()
        
        # Text color
        text_color_layout = QHBoxLayout()
        text_color_layout.addWidget(QLabel("Text Color:"))
        self.text_color_btn = QPushButton("Choose Color")
        self.text_color_btn.clicked.connect(self.choose_text_color)
        text_color_layout.addWidget(self.text_color_btn)
        self.text_color_preview = QLabel("     ")
        self.text_color_preview.setStyleSheet(f"background-color: {self.text_color.name()}; border: 1px solid black;")
        text_color_layout.addWidget(self.text_color_preview)
        color_layout.addLayout(text_color_layout)
        
        # Background color
        bg_color_layout = QHBoxLayout()
        bg_color_layout.addWidget(QLabel("Background:"))
        self.bg_color_btn = QPushButton("Choose Color")
        self.bg_color_btn.clicked.connect(self.choose_bg_color)
        bg_color_layout.addWidget(self.bg_color_btn)
        self.bg_color_preview = QLabel("     ")
        self.bg_color_preview.setStyleSheet(f"background-color: {self.bg_color.name()}; border: 1px solid black;")
        bg_color_layout.addWidget(self.bg_color_preview)
        color_layout.addLayout(bg_color_layout)
        
        color_group.setLayout(color_layout)
        style_layout.addWidget(color_group)
        
        style_layout.addStretch()
        style_tab.setLayout(style_layout)
        self.tabs.addTab(style_tab, "Font & Colors")
        
        # Tab 4: Header Formatting
        header_tab = QWidget()
        header_layout = QVBoxLayout()
        
        header_group = QGroupBox("Header Row Formatting")
        header_form_layout = QVBoxLayout()
        
        self.format_header_cb = QCheckBox("Apply special formatting to header row")
        self.format_header_cb.setChecked(True)
        header_form_layout.addWidget(self.format_header_cb)
        
        # Header colors
        h_text_color_layout = QHBoxLayout()
        h_text_color_layout.addWidget(QLabel("Header Text:"))
        self.header_text_color_btn = QPushButton("Choose Color")
        self.header_text_color_btn.clicked.connect(self.choose_header_text_color)
        h_text_color_layout.addWidget(self.header_text_color_btn)
        self.header_text_preview = QLabel("     ")
        self.header_text_preview.setStyleSheet(f"background-color: {self.header_text_color.name()}; border: 1px solid black;")
        h_text_color_layout.addWidget(self.header_text_preview)
        header_form_layout.addLayout(h_text_color_layout)
        
        h_bg_color_layout = QHBoxLayout()
        h_bg_color_layout.addWidget(QLabel("Header Background:"))
        self.header_bg_color_btn = QPushButton("Choose Color")
        self.header_bg_color_btn.clicked.connect(self.choose_header_bg_color)
        h_bg_color_layout.addWidget(self.header_bg_color_btn)
        self.header_bg_preview = QLabel("     ")
        self.header_bg_preview.setStyleSheet(f"background-color: {self.header_bg_color.name()}; border: 1px solid black;")
        h_bg_color_layout.addWidget(self.header_bg_preview)
        header_form_layout.addLayout(h_bg_color_layout)
        
        self.header_bold_cb = QCheckBox("Bold header text")
        self.header_bold_cb.setChecked(True)
        header_form_layout.addWidget(self.header_bold_cb)
        
        header_group.setLayout(header_form_layout)
        header_layout.addWidget(header_group)
        
        header_layout.addStretch()
        header_tab.setLayout(header_layout)
        self.tabs.addTab(header_tab, "Header")
        
        main_layout.addWidget(self.tabs)
        
        # Action Buttons
        button_layout = QHBoxLayout()
        self.apply_btn = QPushButton("Apply to Column")
        self.apply_btn.clicked.connect(self.apply_to_column)
        button_layout.addWidget(self.apply_btn)
        
        self.apply_all_btn = QPushButton("Apply to All Columns")
        self.apply_all_btn.clicked.connect(self.apply_to_all)
        button_layout.addWidget(self.apply_all_btn)
        
        main_layout.addLayout(button_layout)
        
        # Dialog buttons
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        main_layout.addWidget(button_box)
        
        self.setLayout(main_layout)
    
    def toggle_width_input(self, checked):
        self.width_spin.setEnabled(not checked)
    
    def on_format_type_changed(self, index):
        format_type = self.format_type_combo.currentText()
        
        # Enable/disable relevant controls based on format type
        is_numeric = format_type in ["Number", "Currency", "Accounting", "Percentage", "Scientific"]
        self.decimal_spin.setEnabled(is_numeric)
        self.thousands_cb.setEnabled(is_numeric and format_type != "Scientific")
        
        is_currency = format_type in ["Currency", "Accounting"]
        self.currency_edit.setEnabled(is_currency)
    
    def choose_text_color(self):
        color = QColorDialog.getColor(self.text_color, self, "Choose Text Color")
        if color.isValid():
            self.text_color = color
            self.text_color_preview.setStyleSheet(f"background-color: {color.name()}; border: 1px solid black;")
    
    def choose_bg_color(self):
        color = QColorDialog.getColor(self.bg_color, self, "Choose Background Color")
        if color.isValid():
            self.bg_color = color
            self.bg_color_preview.setStyleSheet(f"background-color: {color.name()}; border: 1px solid black;")
    
    def choose_header_text_color(self):
        color = QColorDialog.getColor(self.header_text_color, self, "Choose Header Text Color")
        if color.isValid():
            self.header_text_color = color
            self.header_text_preview.setStyleSheet(f"background-color: {color.name()}; border: 1px solid black;")
    
    def choose_header_bg_color(self):
        color = QColorDialog.getColor(self.header_bg_color, self, "Choose Header Background Color")
        if color.isValid():
            self.header_bg_color = color
            self.header_bg_preview.setStyleSheet(f"background-color: {color.name()}; border: 1px solid black;")
    
    def on_column_changed(self):
        # Load saved settings for this column if they exist
        column = self.column_combo.currentText()
        if column in self.format_settings:
            settings = self.format_settings[column]
            # Restore settings
            # (Implementation would load saved settings here)
    
    def get_current_settings(self):
        """Get current formatting settings from UI"""
        settings = {
            'width': self.width_spin.value() if not self.auto_width_cb.isChecked() else 'auto',
            'h_align': self.h_align_combo.currentText().lower(),
            'v_align': self.v_align_combo.currentText().lower(),
            'wrap_text': self.wrap_text_cb.isChecked(),
            'format_type': self.format_type_combo.currentText(),
            'decimal_places': self.decimal_spin.value(),
            'currency_symbol': self.currency_edit.text(),
            'use_thousands': self.thousands_cb.isChecked(),
            'font_name': self.font_combo.currentText(),
            'font_size': self.font_size_spin.value(),
            'bold': self.bold_cb.isChecked(),
            'italic': self.italic_cb.isChecked(),
            'underline': self.underline_cb.isChecked(),
            'text_color': self.text_color.name(),
            'bg_color': self.bg_color.name(),
            'format_header': self.format_header_cb.isChecked(),
            'header_text_color': self.header_text_color.name(),
            'header_bg_color': self.header_bg_color.name(),
            'header_bold': self.header_bold_cb.isChecked()
        }
        return settings
    
    def apply_to_column(self):
        """Save settings for current column"""
        column = self.column_combo.currentText()
        self.format_settings[column] = self.get_current_settings()
    
    def apply_to_all(self):
        """Apply current settings to all columns"""
        settings = self.get_current_settings()
        for column in self.columns:
            self.format_settings[str(column)] = settings.copy()
    
    def get_format_settings(self):
        """Return all format settings"""
        return self.format_settings
