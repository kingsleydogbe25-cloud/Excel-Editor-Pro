"""
Data Transformation Module for Excel Editor Pro
Provides comprehensive data transformation tools including:
- Column Operations (Math operations)
- Formula Builder (GUI for Excel formulas)
- Text Tools (Split/Combine, Case, Trim)
- Date Parser
- Data Type Converter
- Fill Down/Fill Series
- Transpose
"""

from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QPushButton, 
                             QLabel, QComboBox, QLineEdit, QMessageBox, QTabWidget,
                             QWidget, QTableWidget, QTableWidgetItem, QCheckBox,
                             QSpinBox, QDoubleSpinBox, QGroupBox, QRadioButton,
                             QButtonGroup, QListWidget, QTextEdit, QDateEdit,
                             QScrollArea, QGridLayout, QFileDialog)
from PyQt5.QtCore import Qt, QDate
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import re


class DataTransformationDialog(QDialog):
    """Main dialog for data transformation operations"""
    
    def __init__(self, df, parent=None):
        super().__init__(parent)
        self.df = df.copy()
        self.original_df = df.copy()
        self.result_df = None
        
        self.setWindowTitle("Data Transformation Tools")
        self.setMinimumSize(900, 700)
        
        self.init_ui()
    
    def init_ui(self):
        layout = QVBoxLayout()
        
        # Create tab widget
        self.tabs = QTabWidget()
        
        # Add all transformation tabs
        self.tabs.addTab(self.create_column_operations_tab(), "📊 Column Math")
        self.tabs.addTab(self.create_formula_builder_tab(), "🔢 Formula Builder")
        self.tabs.addTab(self.create_text_tools_tab(), "📝 Text Tools")
        self.tabs.addTab(self.create_date_parser_tab(), "📅 Date Parser")
        self.tabs.addTab(self.create_type_converter_tab(), "🔄 Type Converter")
        self.tabs.addTab(self.create_fill_tools_tab(), "⬇️ Fill Tools")
        self.tabs.addTab(self.create_transpose_tab(), "↔️ Transpose")
        
        layout.addWidget(self.tabs)
        
        # Buttons
        button_layout = QHBoxLayout()
        
        self.preview_btn = QPushButton("👁 Preview Changes")
        self.preview_btn.clicked.connect(self.preview_changes)
        button_layout.addWidget(self.preview_btn)
        
        self.reset_btn = QPushButton("↺ Reset")
        self.reset_btn.clicked.connect(self.reset_changes)
        button_layout.addWidget(self.reset_btn)
        
        button_layout.addStretch()
        
        self.apply_btn = QPushButton("✓ Apply Transformation")
        self.apply_btn.clicked.connect(self.apply_transformation)
        button_layout.addWidget(self.apply_btn)
        
        self.cancel_btn = QPushButton("✗ Cancel")
        self.cancel_btn.clicked.connect(self.reject)
        button_layout.addWidget(self.cancel_btn)
        
        layout.addLayout(button_layout)
        
        self.setLayout(layout)
    
    # ========== Column Operations Tab ==========
    def create_column_operations_tab(self):
        widget = QWidget()
        layout = QVBoxLayout()
        
        # Operation selection
        op_group = QGroupBox("Math Operations on Columns")
        op_layout = QGridLayout()
        
        op_layout.addWidget(QLabel("Select Column:"), 0, 0)
        self.col_op_column = QComboBox()
        self.col_op_column.addItems(self.df.select_dtypes(include=[np.number]).columns.tolist())
        op_layout.addWidget(self.col_op_column, 0, 1)
        
        op_layout.addWidget(QLabel("Operation:"), 1, 0)
        self.col_op_type = QComboBox()
        self.col_op_type.addItems([
            "Add (+)", "Subtract (-)", "Multiply (×)", "Divide (÷)",
            "Power (^)", "Modulo (%)", "Absolute Value", "Square Root",
            "Round", "Floor", "Ceiling", "Negate", "Percentage of Total"
        ])
        op_layout.addWidget(self.col_op_type, 1, 1)
        
        op_layout.addWidget(QLabel("Value/Column:"), 2, 0)
        self.col_op_value = QLineEdit()
        self.col_op_value.setPlaceholderText("Enter number or column name")
        op_layout.addWidget(self.col_op_value, 2, 1)
        
        op_layout.addWidget(QLabel("Decimals (for rounding):"), 3, 0)
        self.col_op_decimals = QSpinBox()
        self.col_op_decimals.setRange(0, 10)
        self.col_op_decimals.setValue(2)
        op_layout.addWidget(self.col_op_decimals, 3, 1)
        
        op_layout.addWidget(QLabel("Output Column:"), 4, 0)
        self.col_op_output = QLineEdit()
        self.col_op_output.setPlaceholderText("New column name (or leave blank to replace)")
        op_layout.addWidget(self.col_op_output, 4, 1)
        
        self.col_op_btn = QPushButton("Apply Column Operation")
        self.col_op_btn.clicked.connect(self.apply_column_operation)
        op_layout.addWidget(self.col_op_btn, 5, 0, 1, 2)
        
        op_group.setLayout(op_layout)
        layout.addWidget(op_group)
        
        # Multi-column operations
        multi_group = QGroupBox("Multi-Column Operations")
        multi_layout = QGridLayout()
        
        multi_layout.addWidget(QLabel("Operation:"), 0, 0)
        self.multi_op_type = QComboBox()
        self.multi_op_type.addItems(["Sum Columns", "Average Columns", "Min of Columns", 
                                     "Max of Columns", "Product of Columns"])
        multi_layout.addWidget(self.multi_op_type, 0, 1)
        
        multi_layout.addWidget(QLabel("Select Columns:"), 1, 0)
        self.multi_op_columns = QListWidget()
        self.multi_op_columns.setSelectionMode(QListWidget.MultiSelection)
        self.multi_op_columns.addItems(self.df.select_dtypes(include=[np.number]).columns.tolist())
        self.multi_op_columns.setMaximumHeight(100)
        multi_layout.addWidget(self.multi_op_columns, 1, 1)
        
        multi_layout.addWidget(QLabel("Result Column Name:"), 2, 0)
        self.multi_op_result = QLineEdit()
        self.multi_op_result.setPlaceholderText("Name for result column")
        multi_layout.addWidget(self.multi_op_result, 2, 1)
        
        self.multi_op_btn = QPushButton("Apply Multi-Column Operation")
        self.multi_op_btn.clicked.connect(self.apply_multi_column_operation)
        multi_layout.addWidget(self.multi_op_btn, 3, 0, 1, 2)
        
        multi_group.setLayout(multi_layout)
        layout.addWidget(multi_group)
        
        layout.addStretch()
        widget.setLayout(layout)
        return widget
    
    # ========== Formula Builder Tab ==========
    def create_formula_builder_tab(self):
        widget = QWidget()
        layout = QVBoxLayout()
        
        info_label = QLabel("Build Excel-style formulas using column names and functions")
        info_label.setWordWrap(True)
        layout.addWidget(info_label)
        
        # Formula builder
        form_group = QGroupBox("Formula Builder")
        form_layout = QGridLayout()
        
        form_layout.addWidget(QLabel("Available Columns:"), 0, 0)
        self.formula_columns = QListWidget()
        self.formula_columns.addItems(self.df.columns.tolist())
        self.formula_columns.itemDoubleClicked.connect(self.insert_column_to_formula)
        self.formula_columns.setMaximumHeight(120)
        form_layout.addWidget(self.formula_columns, 0, 1, 3, 1)
        
        # Function buttons
        func_widget = QWidget()
        func_layout = QGridLayout()
        func_layout.setSpacing(5)
        
        functions = [
            ("SUM", "SUM()"), ("AVG", "MEAN()"), ("MAX", "MAX()"), ("MIN", "MIN()"),
            ("COUNT", "COUNT()"), ("IF", "IF(condition, true_val, false_val)"),
            ("ROUND", "ROUND(value, decimals)"), ("ABS", "ABS()"),
            ("SQRT", "SQRT()"), ("POWER", "POWER(base, exp)"),
            ("LEN", "LEN()"), ("UPPER", "UPPER()"), ("LOWER", "LOWER()"),
            ("CONCAT", "CONCAT(str1, str2)"), ("TRIM", "TRIM()"),
        ]
        
        row, col = 0, 0
        for label, func in functions:
            btn = QPushButton(label)
            btn.clicked.connect(lambda checked, f=func: self.insert_function(f))
            func_layout.addWidget(btn, row, col)
            col += 1
            if col > 4:
                col = 0
                row += 1
        
        func_widget.setLayout(func_layout)
        form_layout.addWidget(QLabel("Functions:"), 3, 0)
        form_layout.addWidget(func_widget, 3, 1)
        
        form_layout.addWidget(QLabel("Formula:"), 4, 0)
        self.formula_input = QTextEdit()
        self.formula_input.setMaximumHeight(80)
        self.formula_input.setPlaceholderText("Enter formula (e.g., [Column1] * 2 + [Column2])")
        form_layout.addWidget(self.formula_input, 4, 1)
        
        form_layout.addWidget(QLabel("Result Column:"), 5, 0)
        self.formula_result_name = QLineEdit()
        self.formula_result_name.setPlaceholderText("Name for calculated column")
        form_layout.addWidget(self.formula_result_name, 5, 1)
        
        self.formula_apply_btn = QPushButton("Apply Formula")
        self.formula_apply_btn.clicked.connect(self.apply_formula)
        form_layout.addWidget(self.formula_apply_btn, 6, 0, 1, 2)
        
        form_group.setLayout(form_layout)
        layout.addWidget(form_group)
        
        # Examples
        examples_group = QGroupBox("Formula Examples")
        examples_layout = QVBoxLayout()
        examples_text = QLabel(
            "• [Price] * [Quantity] - Calculate total\n"
            "• IF([Sales] > 1000, 'High', 'Low') - Conditional\n"
            "• ROUND([Price] * 1.15, 2) - Add 15% and round\n"
            "• CONCAT([FirstName], ' ', [LastName]) - Combine text\n"
            "• ([Value1] + [Value2]) / 2 - Average of two columns"
        )
        examples_text.setWordWrap(True)
        examples_layout.addWidget(examples_text)
        examples_group.setLayout(examples_layout)
        layout.addWidget(examples_group)
        
        layout.addStretch()
        widget.setLayout(layout)
        return widget
    
    # ========== Text Tools Tab ==========
    def create_text_tools_tab(self):
        widget = QWidget()
        layout = QVBoxLayout()
        
        # Split column
        split_group = QGroupBox("Split Column")
        split_layout = QGridLayout()
        
        split_layout.addWidget(QLabel("Column to Split:"), 0, 0)
        self.split_column = QComboBox()
        self.split_column.addItems(self.df.select_dtypes(include=['object']).columns.tolist())
        split_layout.addWidget(self.split_column, 0, 1)
        
        split_layout.addWidget(QLabel("Split By:"), 1, 0)
        self.split_delimiter = QComboBox()
        self.split_delimiter.addItems(["Space", "Comma", "Tab", "Semicolon", "Custom"])
        self.split_delimiter.currentTextChanged.connect(self.toggle_custom_delimiter)
        split_layout.addWidget(self.split_delimiter, 1, 1)
        
        self.split_custom = QLineEdit()
        self.split_custom.setPlaceholderText("Enter custom delimiter")
        self.split_custom.setEnabled(False)
        split_layout.addWidget(self.split_custom, 1, 2)
        
        split_layout.addWidget(QLabel("Number of Splits:"), 2, 0)
        self.split_max = QSpinBox()
        self.split_max.setRange(1, 20)
        self.split_max.setValue(2)
        split_layout.addWidget(self.split_max, 2, 1)
        
        self.split_btn = QPushButton("Split Column")
        self.split_btn.clicked.connect(self.split_column_action)
        split_layout.addWidget(self.split_btn, 3, 0, 1, 3)
        
        split_group.setLayout(split_layout)
        layout.addWidget(split_group)
        
        # Combine columns
        combine_group = QGroupBox("Combine Columns")
        combine_layout = QGridLayout()
        
        combine_layout.addWidget(QLabel("Select Columns:"), 0, 0)
        self.combine_columns = QListWidget()
        self.combine_columns.setSelectionMode(QListWidget.MultiSelection)
        self.combine_columns.addItems(self.df.columns.tolist())
        self.combine_columns.setMaximumHeight(100)
        combine_layout.addWidget(self.combine_columns, 0, 1, 2, 1)
        
        combine_layout.addWidget(QLabel("Separator:"), 2, 0)
        self.combine_separator = QLineEdit()
        self.combine_separator.setPlaceholderText("e.g., space, comma, dash")
        self.combine_separator.setText(" ")
        combine_layout.addWidget(self.combine_separator, 2, 1)
        
        combine_layout.addWidget(QLabel("Result Column:"), 3, 0)
        self.combine_result = QLineEdit()
        self.combine_result.setPlaceholderText("Name for combined column")
        combine_layout.addWidget(self.combine_result, 3, 1)
        
        self.combine_btn = QPushButton("Combine Columns")
        self.combine_btn.clicked.connect(self.combine_columns_action)
        combine_layout.addWidget(self.combine_btn, 4, 0, 1, 2)
        
        combine_group.setLayout(combine_layout)
        layout.addWidget(combine_group)
        
        # Text transformations
        text_group = QGroupBox("Text Transformations")
        text_layout = QGridLayout()
        
        text_layout.addWidget(QLabel("Column:"), 0, 0)
        self.text_column = QComboBox()
        self.text_column.addItems(self.df.select_dtypes(include=['object']).columns.tolist())
        text_layout.addWidget(self.text_column, 0, 1)
        
        text_layout.addWidget(QLabel("Transformation:"), 1, 0)
        self.text_transform = QComboBox()
        self.text_transform.addItems([
            "UPPERCASE", "lowercase", "Title Case", "Trim Whitespace",
            "Remove Extra Spaces", "Remove Special Chars", "Remove Numbers",
            "Extract Numbers", "Reverse Text", "Remove Duplicates"
        ])
        text_layout.addWidget(self.text_transform, 1, 1)
        
        self.text_inplace = QCheckBox("Apply in place (replace original)")
        self.text_inplace.setChecked(True)
        text_layout.addWidget(self.text_inplace, 2, 0, 1, 2)
        
        self.text_btn = QPushButton("Apply Text Transformation")
        self.text_btn.clicked.connect(self.apply_text_transformation)
        text_layout.addWidget(self.text_btn, 3, 0, 1, 2)
        
        text_group.setLayout(text_layout)
        layout.addWidget(text_group)
        
        layout.addStretch()
        widget.setLayout(layout)
        return widget
    
    # ========== Date Parser Tab ==========
    def create_date_parser_tab(self):
        widget = QWidget()
        layout = QVBoxLayout()
        
        # Parse dates
        parse_group = QGroupBox("Parse Text to Dates")
        parse_layout = QGridLayout()
        
        parse_layout.addWidget(QLabel("Column to Parse:"), 0, 0)
        self.date_column = QComboBox()
        self.date_column.addItems(self.df.columns.tolist())
        parse_layout.addWidget(self.date_column, 0, 1)
        
        parse_layout.addWidget(QLabel("Date Format:"), 1, 0)
        self.date_format = QComboBox()
        self.date_format.addItems([
            "Auto-detect",
            "%Y-%m-%d (2024-01-30)",
            "%m/%d/%Y (01/30/2024)",
            "%d/%m/%Y (30/01/2024)",
            "%Y/%m/%d (2024/01/30)",
            "%d-%m-%Y (30-01-2024)",
            "%b %d, %Y (Jan 30, 2024)",
            "%B %d, %Y (January 30, 2024)",
            "%d %b %Y (30 Jan 2024)",
            "Custom"
        ])
        self.date_format.currentTextChanged.connect(self.toggle_custom_date_format)
        parse_layout.addWidget(self.date_format, 1, 1)
        
        self.date_custom_format = QLineEdit()
        self.date_custom_format.setPlaceholderText("e.g., %Y-%m-%d %H:%M:%S")
        self.date_custom_format.setEnabled(False)
        parse_layout.addWidget(self.date_custom_format, 1, 2)
        
        self.date_parse_btn = QPushButton("Parse to Dates")
        self.date_parse_btn.clicked.connect(self.parse_dates)
        parse_layout.addWidget(self.date_parse_btn, 2, 0, 1, 3)
        
        parse_group.setLayout(parse_layout)
        layout.addWidget(parse_group)
        
        # Extract date components
        extract_group = QGroupBox("Extract Date Components")
        extract_layout = QGridLayout()
        
        extract_layout.addWidget(QLabel("Date Column:"), 0, 0)
        self.extract_date_column = QComboBox()
        self.extract_date_column.addItems(self.df.columns.tolist())
        extract_layout.addWidget(self.extract_date_column, 0, 1)
        
        extract_layout.addWidget(QLabel("Extract:"), 1, 0)
        self.extract_component = QComboBox()
        self.extract_component.addItems([
            "Year", "Month", "Day", "Week of Year", "Day of Week",
            "Quarter", "Day of Year", "Hour", "Minute", "Second"
        ])
        extract_layout.addWidget(self.extract_component, 1, 1)
        
        extract_layout.addWidget(QLabel("Result Column:"), 2, 0)
        self.extract_result_name = QLineEdit()
        self.extract_result_name.setPlaceholderText("Name for extracted component")
        extract_layout.addWidget(self.extract_result_name, 2, 1)
        
        self.extract_btn = QPushButton("Extract Component")
        self.extract_btn.clicked.connect(self.extract_date_component)
        extract_layout.addWidget(self.extract_btn, 3, 0, 1, 2)
        
        extract_group.setLayout(extract_layout)
        layout.addWidget(extract_group)
        
        # Date calculations
        calc_group = QGroupBox("Date Calculations")
        calc_layout = QGridLayout()
        
        calc_layout.addWidget(QLabel("Date Column:"), 0, 0)
        self.calc_date_column = QComboBox()
        self.calc_date_column.addItems(self.df.columns.tolist())
        calc_layout.addWidget(self.calc_date_column, 0, 1)
        
        calc_layout.addWidget(QLabel("Operation:"), 1, 0)
        self.date_calc_op = QComboBox()
        self.date_calc_op.addItems(["Add Days", "Subtract Days", "Days Until Today", 
                                    "Days Since Date", "Age in Years"])
        calc_layout.addWidget(self.date_calc_op, 1, 1)
        
        calc_layout.addWidget(QLabel("Value:"), 2, 0)
        self.date_calc_value = QSpinBox()
        self.date_calc_value.setRange(-10000, 10000)
        self.date_calc_value.setValue(0)
        calc_layout.addWidget(self.date_calc_value, 2, 1)
        
        calc_layout.addWidget(QLabel("Result Column:"), 3, 0)
        self.calc_result_name = QLineEdit()
        self.calc_result_name.setPlaceholderText("Name for result")
        calc_layout.addWidget(self.calc_result_name, 3, 1)
        
        self.date_calc_btn = QPushButton("Calculate")
        self.date_calc_btn.clicked.connect(self.calculate_dates)
        calc_layout.addWidget(self.date_calc_btn, 4, 0, 1, 2)
        
        calc_group.setLayout(calc_layout)
        layout.addWidget(calc_group)
        
        layout.addStretch()
        widget.setLayout(layout)
        return widget
    
    # ========== Type Converter Tab ==========
    def create_type_converter_tab(self):
        widget = QWidget()
        layout = QVBoxLayout()
        
        info = QLabel("Convert data types of columns (batch processing supported)")
        layout.addWidget(info)
        
        convert_group = QGroupBox("Data Type Conversion")
        convert_layout = QGridLayout()
        
        convert_layout.addWidget(QLabel("Select Columns:"), 0, 0)
        self.convert_columns = QListWidget()
        self.convert_columns.setSelectionMode(QListWidget.MultiSelection)
        self.convert_columns.addItems(self.df.columns.tolist())
        self.convert_columns.setMaximumHeight(150)
        convert_layout.addWidget(self.convert_columns, 0, 1, 3, 1)
        
        convert_layout.addWidget(QLabel("Convert To:"), 3, 0)
        self.convert_type = QComboBox()
        self.convert_type.addItems([
            "Integer", "Float", "String (Text)", "Boolean", 
            "Date/Time", "Category"
        ])
        convert_layout.addWidget(self.convert_type, 3, 1)
        
        self.convert_errors = QComboBox()
        self.convert_errors.addItems(["Coerce (set invalid to NaN)", "Ignore (keep original)", "Raise error"])
        convert_layout.addWidget(QLabel("On Error:"), 4, 0)
        convert_layout.addWidget(self.convert_errors, 4, 1)
        
        self.convert_btn = QPushButton("Convert Data Types")
        self.convert_btn.clicked.connect(self.convert_data_types)
        convert_layout.addWidget(self.convert_btn, 5, 0, 1, 2)
        
        convert_group.setLayout(convert_layout)
        layout.addWidget(convert_group)
        
        # Current types display
        types_group = QGroupBox("Current Column Types")
        types_layout = QVBoxLayout()
        
        self.types_display = QTextEdit()
        self.types_display.setReadOnly(True)
        self.types_display.setMaximumHeight(150)
        self.update_types_display()
        types_layout.addWidget(self.types_display)
        
        refresh_btn = QPushButton("↻ Refresh Types")
        refresh_btn.clicked.connect(self.update_types_display)
        types_layout.addWidget(refresh_btn)
        
        types_group.setLayout(types_layout)
        layout.addWidget(types_group)
        
        layout.addStretch()
        widget.setLayout(layout)
        return widget
    
    # ========== Fill Tools Tab ==========
    def create_fill_tools_tab(self):
        widget = QWidget()
        layout = QVBoxLayout()
        
        # Fill Down
        fill_down_group = QGroupBox("Fill Down")
        fill_down_layout = QGridLayout()
        
        fill_down_layout.addWidget(QLabel("Column:"), 0, 0)
        self.fill_down_column = QComboBox()
        self.fill_down_column.addItems(self.df.columns.tolist())
        fill_down_layout.addWidget(self.fill_down_column, 0, 1)
        
        self.fill_down_btn = QPushButton("Fill Down (Copy first value to all empty cells)")
        self.fill_down_btn.clicked.connect(self.fill_down)
        fill_down_layout.addWidget(self.fill_down_btn, 1, 0, 1, 2)
        
        fill_down_group.setLayout(fill_down_layout)
        layout.addWidget(fill_down_group)
        
        # Fill Series
        series_group = QGroupBox("Fill Series (Auto-Fill Patterns)")
        series_layout = QGridLayout()
        
        series_layout.addWidget(QLabel("Column:"), 0, 0)
        self.series_column = QComboBox()
        self.series_column.addItems(self.df.columns.tolist())
        series_layout.addWidget(self.series_column, 0, 1)
        
        series_layout.addWidget(QLabel("Series Type:"), 1, 0)
        self.series_type = QComboBox()
        self.series_type.addItems([
            "Number Sequence (1, 2, 3...)",
            "Date Sequence (Daily)",
            "Date Sequence (Weekly)", 
            "Date Sequence (Monthly)",
            "Custom Pattern"
        ])
        self.series_type.currentTextChanged.connect(self.toggle_series_options)
        series_layout.addWidget(self.series_type, 1, 1)
        
        series_layout.addWidget(QLabel("Start Value:"), 2, 0)
        self.series_start = QLineEdit()
        self.series_start.setPlaceholderText("1")
        series_layout.addWidget(self.series_start, 2, 1)
        
        series_layout.addWidget(QLabel("Step/Increment:"), 3, 0)
        self.series_step = QLineEdit()
        self.series_step.setPlaceholderText("1")
        series_layout.addWidget(self.series_step, 3, 1)
        
        series_layout.addWidget(QLabel("Apply To:"), 4, 0)
        self.series_range = QComboBox()
        self.series_range.addItems(["All Rows", "Empty Cells Only", "Selected Range"])
        series_layout.addWidget(self.series_range, 4, 1)
        
        self.series_btn = QPushButton("Fill Series")
        self.series_btn.clicked.connect(self.fill_series)
        series_layout.addWidget(self.series_btn, 5, 0, 1, 2)
        
        series_group.setLayout(series_layout)
        layout.addWidget(series_group)
        
        # Forward Fill / Backward Fill
        ffill_group = QGroupBox("Forward/Backward Fill")
        ffill_layout = QGridLayout()
        
        ffill_layout.addWidget(QLabel("Column:"), 0, 0)
        self.ffill_column = QComboBox()
        self.ffill_column.addItems(self.df.columns.tolist())
        ffill_layout.addWidget(self.ffill_column, 0, 1)
        
        ffill_layout.addWidget(QLabel("Method:"), 1, 0)
        self.ffill_method = QComboBox()
        self.ffill_method.addItems(["Forward Fill (use previous)", "Backward Fill (use next)"])
        ffill_layout.addWidget(self.ffill_method, 1, 1)
        
        self.ffill_btn = QPushButton("Apply Fill")
        self.ffill_btn.clicked.connect(self.forward_backward_fill)
        ffill_layout.addWidget(self.ffill_btn, 2, 0, 1, 2)
        
        ffill_group.setLayout(ffill_layout)
        layout.addWidget(ffill_group)
        
        layout.addStretch()
        widget.setLayout(layout)
        return widget
    
    # ========== Transpose Tab ==========
    def create_transpose_tab(self):
        widget = QWidget()
        layout = QVBoxLayout()
        
        info = QLabel("⚠️ Transpose will flip rows and columns. This creates a new dataframe structure.")
        info.setWordWrap(True)
        info.setStyleSheet("color: orange; font-weight: bold; padding: 10px;")
        layout.addWidget(info)
        
        transpose_group = QGroupBox("Transpose Options")
        transpose_layout = QVBoxLayout()
        
        self.transpose_keep_index = QCheckBox("Keep current index as first column")
        self.transpose_keep_index.setChecked(True)
        transpose_layout.addWidget(self.transpose_keep_index)
        
        self.transpose_keep_columns = QCheckBox("Use first row as new column headers")
        self.transpose_keep_columns.setChecked(True)
        transpose_layout.addWidget(self.transpose_keep_columns)
        
        # Preview
        preview_label = QLabel("\nCurrent Shape: {} rows × {} columns\nAfter Transpose: {} rows × {} columns".format(
            len(self.df), len(self.df.columns), len(self.df.columns), len(self.df)
        ))
        preview_label.setWordWrap(True)
        transpose_layout.addWidget(preview_label)
        
        self.transpose_btn = QPushButton("🔄 Transpose Data")
        self.transpose_btn.clicked.connect(self.transpose_data)
        transpose_layout.addWidget(self.transpose_btn)
        
        transpose_group.setLayout(transpose_layout)
        layout.addWidget(transpose_group)
        
        # Example visualization
        example_group = QGroupBox("Transpose Example")
        example_layout = QHBoxLayout()
        
        before = QLabel("Before:\n┌────────┬────────┐\n│ Name   │ Value  │\n├────────┼────────┤\n│ Alice  │ 100    │\n│ Bob    │ 200    │\n└────────┴────────┘")
        before.setStyleSheet("font-family: monospace;")
        example_layout.addWidget(before)
        
        arrow = QLabel("  →  ")
        arrow.setStyleSheet("font-size: 24px;")
        example_layout.addWidget(arrow)
        
        after = QLabel("After:\n┌────────┬────────┬────────┐\n│        │ Alice  │ Bob    │\n├────────┼────────┼────────┤\n│ Value  │ 100    │ 200    │\n└────────┴────────┴────────┘")
        after.setStyleSheet("font-family: monospace;")
        example_layout.addWidget(after)
        
        example_group.setLayout(example_layout)
        layout.addWidget(example_group)
        
        layout.addStretch()
        widget.setLayout(layout)
        return widget
    
    # ========== Helper Functions ==========
    
    def toggle_custom_delimiter(self, text):
        self.split_custom.setEnabled(text == "Custom")
    
    def toggle_custom_date_format(self, text):
        self.date_custom_format.setEnabled(text == "Custom")
    
    def toggle_series_options(self, text):
        # Enable/disable options based on series type
        pass
    
    def insert_column_to_formula(self, item):
        """Insert column name into formula with brackets"""
        col_name = item.text()
        self.formula_input.insertPlainText(f"[{col_name}]")
    
    def insert_function(self, func):
        """Insert function into formula"""
        self.formula_input.insertPlainText(func)
    
    def update_types_display(self):
        """Display current data types of all columns"""
        types_info = "Column Types:\n" + "="*40 + "\n"
        for col in self.df.columns:
            dtype = str(self.df[col].dtype)
            types_info += f"{col}: {dtype}\n"
        self.types_display.setText(types_info)
    
    # ========== Action Functions ==========
    
    def apply_column_operation(self):
        """Apply math operation to a column"""
        try:
            column = self.col_op_column.currentText()
            operation = self.col_op_type.currentText()
            value_str = self.col_op_value.text()
            output_col = self.col_op_output.text() or column
            decimals = self.col_op_decimals.value()
            
            if column not in self.df.columns:
                QMessageBox.warning(self, "Error", "Column not found")
                return
            
            # Parse value (could be a number or column name)
            if value_str in self.df.columns:
                value = self.df[value_str]
            else:
                try:
                    value = float(value_str) if value_str else 1
                except:
                    value = 1
            
            # Apply operation
            if operation == "Add (+)":
                result = self.df[column] + value
            elif operation == "Subtract (-)":
                result = self.df[column] - value
            elif operation == "Multiply (×)":
                result = self.df[column] * value
            elif operation == "Divide (÷)":
                result = self.df[column] / value
            elif operation == "Power (^)":
                result = self.df[column] ** value
            elif operation == "Modulo (%)":
                result = self.df[column] % value
            elif operation == "Absolute Value":
                result = self.df[column].abs()
            elif operation == "Square Root":
                result = np.sqrt(self.df[column])
            elif operation == "Round":
                result = self.df[column].round(decimals)
            elif operation == "Floor":
                result = np.floor(self.df[column])
            elif operation == "Ceiling":
                result = np.ceil(self.df[column])
            elif operation == "Negate":
                result = -self.df[column]
            elif operation == "Percentage of Total":
                result = (self.df[column] / self.df[column].sum()) * 100
            
            self.df[output_col] = result
            QMessageBox.information(self, "Success", f"Operation applied to column '{output_col}'")
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to apply operation: {str(e)}")
    
    def apply_multi_column_operation(self):
        """Apply operation across multiple columns"""
        try:
            operation = self.multi_op_type.currentText()
            selected = [item.text() for item in self.multi_op_columns.selectedItems()]
            result_name = self.multi_op_result.text()
            
            if not selected:
                QMessageBox.warning(self, "Error", "Please select at least one column")
                return
            
            if not result_name:
                QMessageBox.warning(self, "Error", "Please enter a result column name")
                return
            
            # Apply operation
            if operation == "Sum Columns":
                self.df[result_name] = self.df[selected].sum(axis=1)
            elif operation == "Average Columns":
                self.df[result_name] = self.df[selected].mean(axis=1)
            elif operation == "Min of Columns":
                self.df[result_name] = self.df[selected].min(axis=1)
            elif operation == "Max of Columns":
                self.df[result_name] = self.df[selected].max(axis=1)
            elif operation == "Product of Columns":
                self.df[result_name] = self.df[selected].prod(axis=1)
            
            QMessageBox.information(self, "Success", f"Multi-column operation completed: '{result_name}'")
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed: {str(e)}")
    
    def apply_formula(self):
        """Apply custom formula"""
        try:
            formula = self.formula_input.toPlainText().strip()
            result_name = self.formula_result_name.text()
            
            if not formula or not result_name:
                QMessageBox.warning(self, "Error", "Please enter both formula and result column name")
                return
            
            # Replace column names in brackets with actual column references
            eval_formula = formula
            for col in self.df.columns:
                eval_formula = eval_formula.replace(f"[{col}]", f"self.df['{col}']")
            
            # Replace common Excel functions with pandas/numpy equivalents
            replacements = {
                'SUM(': 'sum(',
                'AVG(': 'mean(',
                'MEAN(': 'mean(',
                'MAX(': 'max(',
                'MIN(': 'min(',
                'COUNT(': 'count(',
                'ABS(': 'abs(',
                'SQRT(': 'np.sqrt(',
                'POWER(': 'np.power(',
                'ROUND(': 'round(',
                'UPPER(': 'str.upper(',
                'LOWER(': 'str.lower(',
                'LEN(': 'len(',
                'TRIM(': 'str.strip(',
                'CONCAT(': 'str.cat(',
            }
            
            for old, new in replacements.items():
                eval_formula = eval_formula.replace(old, new)
            
            # Evaluate formula
            result = eval(eval_formula)
            self.df[result_name] = result
            
            QMessageBox.information(self, "Success", f"Formula applied successfully: '{result_name}'")
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Formula error: {str(e)}\n\nPlease check your formula syntax.")
    
    def split_column_action(self):
        """Split a column by delimiter"""
        try:
            column = self.split_column.currentText()
            delimiter_type = self.split_delimiter.currentText()
            max_splits = self.split_max.value()
            
            # Get delimiter
            delimiters = {
                "Space": " ",
                "Comma": ",",
                "Tab": "\t",
                "Semicolon": ";",
                "Custom": self.split_custom.text()
            }
            delimiter = delimiters.get(delimiter_type, " ")
            
            # Split column
            split_data = self.df[column].str.split(delimiter, n=max_splits, expand=True)
            
            # Add new columns
            for i in range(split_data.shape[1]):
                new_col = f"{column}_Part{i+1}"
                self.df[new_col] = split_data[i]
            
            QMessageBox.information(self, "Success", 
                                  f"Column '{column}' split into {split_data.shape[1]} parts")
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to split column: {str(e)}")
    
    def combine_columns_action(self):
        """Combine multiple columns"""
        try:
            selected = [item.text() for item in self.combine_columns.selectedItems()]
            separator = self.combine_separator.text()
            result_name = self.combine_result.text()
            
            if len(selected) < 2:
                QMessageBox.warning(self, "Error", "Please select at least 2 columns to combine")
                return
            
            if not result_name:
                QMessageBox.warning(self, "Error", "Please enter a result column name")
                return
            
            # Combine columns
            self.df[result_name] = self.df[selected].apply(
                lambda x: separator.join(x.astype(str)), axis=1
            )
            
            QMessageBox.information(self, "Success", f"Columns combined into '{result_name}'")
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to combine columns: {str(e)}")
    
    def apply_text_transformation(self):
        """Apply text transformation to column"""
        try:
            column = self.text_column.currentText()
            transform = self.text_transform.currentText()
            in_place = self.text_inplace.isChecked()
            
            target_col = column if in_place else f"{column}_Transformed"
            
            if transform == "UPPERCASE":
                self.df[target_col] = self.df[column].str.upper()
            elif transform == "lowercase":
                self.df[target_col] = self.df[column].str.lower()
            elif transform == "Title Case":
                self.df[target_col] = self.df[column].str.title()
            elif transform == "Trim Whitespace":
                self.df[target_col] = self.df[column].str.strip()
            elif transform == "Remove Extra Spaces":
                self.df[target_col] = self.df[column].str.replace(r'\s+', ' ', regex=True)
            elif transform == "Remove Special Chars":
                self.df[target_col] = self.df[column].str.replace(r'[^a-zA-Z0-9\s]', '', regex=True)
            elif transform == "Remove Numbers":
                self.df[target_col] = self.df[column].str.replace(r'\d+', '', regex=True)
            elif transform == "Extract Numbers":
                self.df[target_col] = self.df[column].str.extract(r'(\d+)', expand=False)
            elif transform == "Reverse Text":
                self.df[target_col] = self.df[column].str[::-1]
            elif transform == "Remove Duplicates":
                self.df[target_col] = self.df[column].drop_duplicates()
            
            QMessageBox.information(self, "Success", "Text transformation applied")
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed: {str(e)}")
    
    def parse_dates(self):
        """Parse text to dates"""
        try:
            column = self.date_column.currentText()
            format_type = self.date_format.currentText()
            
            # Get format string
            formats = {
                "Auto-detect": None,
                "%Y-%m-%d (2024-01-30)": "%Y-%m-%d",
                "%m/%d/%Y (01/30/2024)": "%m/%d/%Y",
                "%d/%m/%Y (30/01/2024)": "%d/%m/%Y",
                "%Y/%m/%d (2024/01/30)": "%Y/%m/%d",
                "%d-%m-%Y (30-01-2024)": "%d-%m-%Y",
                "%b %d, %Y (Jan 30, 2024)": "%b %d, %Y",
                "%B %d, %Y (January 30, 2024)": "%B %d, %Y",
                "%d %b %Y (30 Jan 2024)": "%d %b %Y",
                "Custom": self.date_custom_format.text()
            }
            
            date_format = formats.get(format_type)
            
            if date_format:
                self.df[column] = pd.to_datetime(self.df[column], format=date_format, errors='coerce')
            else:
                self.df[column] = pd.to_datetime(self.df[column], infer_datetime_format=True, errors='coerce')
            
            QMessageBox.information(self, "Success", f"Dates parsed in column '{column}'")
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to parse dates: {str(e)}")
    
    def extract_date_component(self):
        """Extract component from date column"""
        try:
            column = self.extract_date_column.currentText()
            component = self.extract_component.currentText()
            result_name = self.extract_result_name.text()
            
            if not result_name:
                result_name = f"{column}_{component}"
            
            # Ensure column is datetime
            if not pd.api.types.is_datetime64_any_dtype(self.df[column]):
                self.df[column] = pd.to_datetime(self.df[column], errors='coerce')
            
            # Extract component
            if component == "Year":
                self.df[result_name] = self.df[column].dt.year
            elif component == "Month":
                self.df[result_name] = self.df[column].dt.month
            elif component == "Day":
                self.df[result_name] = self.df[column].dt.day
            elif component == "Week of Year":
                self.df[result_name] = self.df[column].dt.isocalendar().week
            elif component == "Day of Week":
                self.df[result_name] = self.df[column].dt.dayofweek
            elif component == "Quarter":
                self.df[result_name] = self.df[column].dt.quarter
            elif component == "Day of Year":
                self.df[result_name] = self.df[column].dt.dayofyear
            elif component == "Hour":
                self.df[result_name] = self.df[column].dt.hour
            elif component == "Minute":
                self.df[result_name] = self.df[column].dt.minute
            elif component == "Second":
                self.df[result_name] = self.df[column].dt.second
            
            QMessageBox.information(self, "Success", f"Extracted '{component}' to '{result_name}'")
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed: {str(e)}")
    
    def calculate_dates(self):
        """Perform date calculations"""
        try:
            column = self.calc_date_column.currentText()
            operation = self.date_calc_op.currentText()
            value = self.date_calc_value.value()
            result_name = self.calc_result_name.text() or f"{column}_Calculated"
            
            # Ensure column is datetime
            if not pd.api.types.is_datetime64_any_dtype(self.df[column]):
                self.df[column] = pd.to_datetime(self.df[column], errors='coerce')
            
            if operation == "Add Days":
                self.df[result_name] = self.df[column] + pd.Timedelta(days=value)
            elif operation == "Subtract Days":
                self.df[result_name] = self.df[column] - pd.Timedelta(days=value)
            elif operation == "Days Until Today":
                today = pd.Timestamp.now()
                self.df[result_name] = (today - self.df[column]).dt.days
            elif operation == "Days Since Date":
                self.df[result_name] = (self.df[column] - pd.Timestamp.now()).dt.days
            elif operation == "Age in Years":
                today = pd.Timestamp.now()
                self.df[result_name] = ((today - self.df[column]).dt.days / 365.25).astype(int)
            
            QMessageBox.information(self, "Success", f"Date calculation completed: '{result_name}'")
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed: {str(e)}")
    
    def convert_data_types(self):
        """Convert column data types"""
        try:
            selected = [item.text() for item in self.convert_columns.selectedItems()]
            target_type = self.convert_type.currentText()
            error_handling = self.convert_errors.currentText()
            
            if not selected:
                QMessageBox.warning(self, "Error", "Please select columns to convert")
                return
            
            errors_param = 'coerce' if 'Coerce' in error_handling else 'ignore'
            
            for col in selected:
                try:
                    if target_type == "Integer":
                        self.df[col] = pd.to_numeric(self.df[col], errors=errors_param).astype('Int64')
                    elif target_type == "Float":
                        self.df[col] = pd.to_numeric(self.df[col], errors=errors_param)
                    elif target_type == "String (Text)":
                        self.df[col] = self.df[col].astype(str)
                    elif target_type == "Boolean":
                        self.df[col] = self.df[col].astype(bool)
                    elif target_type == "Date/Time":
                        self.df[col] = pd.to_datetime(self.df[col], errors=errors_param)
                    elif target_type == "Category":
                        self.df[col] = self.df[col].astype('category')
                except Exception as e:
                    if 'Raise' in error_handling:
                        raise e
            
            self.update_types_display()
            QMessageBox.information(self, "Success", f"Converted {len(selected)} column(s) to {target_type}")
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Conversion failed: {str(e)}")
    
    def fill_down(self):
        """Fill down first non-null value"""
        try:
            column = self.fill_down_column.currentText()
            self.df[column] = self.df[column].fillna(method='ffill')
            QMessageBox.information(self, "Success", f"Filled down column '{column}'")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed: {str(e)}")
    
    def fill_series(self):
        """Fill column with series pattern"""
        try:
            column = self.series_column.currentText()
            series_type = self.series_type.currentText()
            start_val = self.series_start.text()
            step_val = self.series_step.text()
            
            num_rows = len(self.df)
            
            if "Number" in series_type:
                start = float(start_val) if start_val else 1
                step = float(step_val) if step_val else 1
                series = [start + i * step for i in range(num_rows)]
                self.df[column] = series
                
            elif "Daily" in series_type:
                start = pd.to_datetime(start_val) if start_val else pd.Timestamp.now()
                series = pd.date_range(start=start, periods=num_rows, freq='D')
                self.df[column] = series
                
            elif "Weekly" in series_type:
                start = pd.to_datetime(start_val) if start_val else pd.Timestamp.now()
                series = pd.date_range(start=start, periods=num_rows, freq='W')
                self.df[column] = series
                
            elif "Monthly" in series_type:
                start = pd.to_datetime(start_val) if start_val else pd.Timestamp.now()
                series = pd.date_range(start=start, periods=num_rows, freq='MS')
                self.df[column] = series
            
            QMessageBox.information(self, "Success", f"Series filled in column '{column}'")
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed: {str(e)}")
    
    def forward_backward_fill(self):
        """Apply forward or backward fill"""
        try:
            column = self.ffill_column.currentText()
            method = self.ffill_method.currentText()
            
            if "Forward" in method:
                self.df[column] = self.df[column].fillna(method='ffill')
            else:
                self.df[column] = self.df[column].fillna(method='bfill')
            
            QMessageBox.information(self, "Success", f"Fill applied to '{column}'")
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed: {str(e)}")
    
    def transpose_data(self):
        """Transpose the dataframe"""
        try:
            reply = QMessageBox.question(
                self, 
                "Confirm Transpose",
                "This will completely restructure your data. Continue?",
                QMessageBox.Yes | QMessageBox.No
            )
            
            if reply == QMessageBox.Yes:
                self.df = self.df.T
                
                if not self.transpose_keep_columns.isChecked():
                    self.df.columns = [f"Col_{i+1}" for i in range(len(self.df.columns))]
                
                if not self.transpose_keep_index.isChecked():
                    self.df.reset_index(drop=True, inplace=True)
                
                QMessageBox.information(self, "Success", 
                                      f"Data transposed! New shape: {len(self.df)} rows × {len(self.df.columns)} columns")
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Transpose failed: {str(e)}")
    
    def preview_changes(self):
        """Preview the transformed data"""
        try:
            preview_dialog = QDialog(self)
            preview_dialog.setWindowTitle("Preview Transformed Data")
            preview_dialog.setMinimumSize(800, 600)
            
            layout = QVBoxLayout()
            
            info = QLabel(f"Showing first 100 rows of transformed data\nTotal rows: {len(self.df)}, Total columns: {len(self.df.columns)}")
            layout.addWidget(info)
            
            table = QTableWidget()
            table.setRowCount(min(100, len(self.df)))
            table.setColumnCount(len(self.df.columns))
            table.setHorizontalHeaderLabels(self.df.columns.tolist())
            
            for i in range(min(100, len(self.df))):
                for j, col in enumerate(self.df.columns):
                    item = QTableWidgetItem(str(self.df.iloc[i, j]))
                    table.setItem(i, j, item)
            
            table.resizeColumnsToContents()
            layout.addWidget(table)
            
            close_btn = QPushButton("Close Preview")
            close_btn.clicked.connect(preview_dialog.accept)
            layout.addWidget(close_btn)
            
            preview_dialog.setLayout(layout)
            preview_dialog.exec_()
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Preview failed: {str(e)}")
    
    def reset_changes(self):
        """Reset to original data"""
        self.df = self.original_df.copy()
        QMessageBox.information(self, "Reset", "All changes have been reset to original data")
    
    def apply_transformation(self):
        """Apply all transformations and close dialog"""
        self.result_df = self.df.copy()
        self.accept()
    
    def get_result(self):
        """Return the transformed dataframe"""
        return self.result_df
