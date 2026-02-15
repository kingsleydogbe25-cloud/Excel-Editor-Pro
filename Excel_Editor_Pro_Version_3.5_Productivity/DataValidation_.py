from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QPushButton, 
                             QLabel, QComboBox, QWidget, QMessageBox, QGroupBox,
                             QLineEdit, QItemDelegate, QTableWidget)
from PyQt5.QtCore import Qt, QRegExp
from PyQt5.QtGui import QColor, QBrush, QRegExpValidator

class ValidationRule:
    TYPE_ANY = "Any Value"
    TYPE_NUMBER = "Number"
    TYPE_TEXT = "Text Length"
    TYPE_LIST = "List"
    TYPE_REGEX = "Regex / Text Content"

    def __init__(self, rule_type=TYPE_ANY, **kwargs):
        self.type = rule_type
        self.params = kwargs

    def validate(self, value):
        if self.type == self.TYPE_ANY:
            return True, ""
            
        try:
            if self.type == self.TYPE_NUMBER:
                num = float(value)
                op = self.params.get('operator', '>')
                threshold = float(self.params.get('value', 0))
                
                if op == '>' and not (num > threshold): return False, f"Value must be > {threshold}"
                if op == '<' and not (num < threshold): return False, f"Value must be < {threshold}"
                if op == '>=' and not (num >= threshold): return False, f"Value must be >= {threshold}"
                if op == '<=' and not (num <= threshold): return False, f"Value must be <= {threshold}"
                if op == '=' and not (num == threshold): return False, f"Value must be = {threshold}"
                
            elif self.type == self.TYPE_TEXT:
                length = len(str(value))
                op = self.params.get('operator', '>')
                threshold = int(self.params.get('value', 0))
                
                if op == '>' and not (length > threshold): return False, f"Text length must be > {threshold}"
                if op == '<' and not (length < threshold): return False, f"Text length must be < {threshold}"
                
            elif self.type == self.TYPE_LIST:
                options = self.params.get('options', [])
                if str(value) not in options: return False, f"Value must be one of: {', '.join(options)}"
                
            elif self.type == self.TYPE_REGEX:
                pattern = self.params.get('pattern', '')
                if pattern == "Contains":
                    if self.params.get('value', '') not in str(value):
                        return False, f"Text must contain '{self.params.get('value', '')}'"
                elif pattern == "Starts With":
                    if not str(value).startswith(self.params.get('value', '')):
                         return False, f"Text must start with '{self.params.get('value', '')}'"
                
        except ValueError:
            return False, "Invalid data type"
        except Exception as e:
            return False, str(e)
            
        return True, ""

class ValidationManager:
    def __init__(self):
        # Dictionary mapping column index (int) to ValidationRule
        self.rules = {}

    def add_rule(self, col_index, rule):
        self.rules[col_index] = rule

    def get_rule(self, col_index):
        return self.rules.get(col_index)
        
    def remove_rule(self, col_index):
        if col_index in self.rules:
            del self.rules[col_index]

    def validate_cell(self, col_index, value):
        rule = self.rules.get(col_index)
        if not rule:
            return True, ""
        return rule.validate(value)

class DropdownDelegate(QItemDelegate):
    def __init__(self, parent=None, items=None):
        super().__init__(parent)
        self.items = items or []

    def createEditor(self, parent, option, index):
        editor = QComboBox(parent)
        editor.addItems(self.items)
        return editor

    def setEditorData(self, editor, index):
        value = index.model().data(index, Qt.EditRole)
        if value:
            idx = editor.findText(str(value))
            if idx >= 0:
                editor.setCurrentIndex(idx)

    def setModelData(self, editor, model, index):
        value = editor.currentText()
        model.setData(index, value, Qt.EditRole)

class ValidationDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Data Validation Rule")
        self.resize(400, 300)
        self.rule = None
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        
        # Rule Type
        layout.addWidget(QLabel("Validation Type:"))
        self.type_combo = QComboBox()
        self.type_combo.addItems([
            ValidationRule.TYPE_ANY,
            ValidationRule.TYPE_NUMBER,
            ValidationRule.TYPE_TEXT,
            ValidationRule.TYPE_LIST,
            ValidationRule.TYPE_REGEX
        ])
        self.type_combo.currentTextChanged.connect(self.update_ui)
        layout.addWidget(self.type_combo)

        # Params Area
        self.params_group = QGroupBox("Rule Parameters")
        self.params_layout = QVBoxLayout()
        self.params_group.setLayout(self.params_layout)
        layout.addWidget(self.params_group)
        
        # Initial fields (will be cleared/repopulated)
        self.operator_combo = QComboBox()
        self.value_input = QLineEdit()
        self.list_input = QLineEdit()
        
        self.update_ui()
        
        # Buttons
        btn_layout = QHBoxLayout()
        ok_btn = QPushButton("OK")
        ok_btn.clicked.connect(self.accept_rule)
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        
        btn_layout.addStretch()
        btn_layout.addWidget(ok_btn)
        btn_layout.addWidget(cancel_btn)
        layout.addLayout(btn_layout)
        
        self.setLayout(layout)

    def update_ui(self):
        # Clear previous layout
        while self.params_layout.count():
            item = self.params_layout.takeAt(0)
            widget = item.widget()
            if widget: widget.deleteLater()
            
        rule_type = self.type_combo.currentText()
        
        if rule_type == ValidationRule.TYPE_NUMBER or rule_type == ValidationRule.TYPE_TEXT:
            self.params_layout.addWidget(QLabel("Operator:"))
            self.operator_combo = QComboBox()
            self.operator_combo.addItems([">", "<", ">=", "<=", "="])
            self.params_layout.addWidget(self.operator_combo)
            
            self.params_layout.addWidget(QLabel("Value / Length:"))
            self.value_input = QLineEdit()
            self.params_layout.addWidget(self.value_input)
            
        elif rule_type == ValidationRule.TYPE_LIST:
            self.params_layout.addWidget(QLabel("Allowed Values (comma separated):"))
            self.list_input = QLineEdit()
            self.list_input.setPlaceholderText("Apple, Banana, Orange")
            self.params_layout.addWidget(self.list_input)
            
        elif rule_type == ValidationRule.TYPE_REGEX:
            self.params_layout.addWidget(QLabel("Condition:"))
            self.operator_combo = QComboBox()
            self.operator_combo.addItems(["Contains", "Starts With"])
            self.params_layout.addWidget(self.operator_combo)
            
            self.params_layout.addWidget(QLabel("Text:"))
            self.value_input = QLineEdit()
            self.params_layout.addWidget(self.value_input)

    def accept_rule(self):
        rule_type = self.type_combo.currentText()
        params = {}
        
        try:
            if rule_type == ValidationRule.TYPE_NUMBER:
                params['operator'] = self.operator_combo.currentText()
                params['value'] = float(self.value_input.text())
            elif rule_type == ValidationRule.TYPE_TEXT:
                params['operator'] = self.operator_combo.currentText()
                params['value'] = int(self.value_input.text())
            elif rule_type == ValidationRule.TYPE_LIST:
                text = self.list_input.text()
                params['options'] = [x.strip() for x in text.split(',') if x.strip()]
            elif rule_type == ValidationRule.TYPE_REGEX:
                params['pattern'] = self.operator_combo.currentText()
                params['value'] = self.value_input.text()
                
            self.rule = ValidationRule(rule_type, **params)
            self.accept()
        except ValueError:
            QMessageBox.warning(self, "Invalid Input", "Please enter valid numeric values for this rule type.")

    def get_rule(self):
        return self.rule
