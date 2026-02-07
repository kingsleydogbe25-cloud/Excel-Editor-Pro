from PyQt5.QtWidgets import(QDialog, QVBoxLayout, QLabel, QGroupBox, QWidget, QHBoxLayout,
                            QPushButton, QCheckBox, QDialogButtonBox,QTabWidget, QFormLayout, 
                            QSpinBox, QColorDialog)
from PyQt5.QtGui import QColor


class SettingsDialog(QDialog):
    """Dialog for application settings including theme and auto-save"""
    def __init__(self, current_bg, current_accent, auto_save_enabled, auto_save_interval, keep_versions=10, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Application Settings")
        self.setModal(True)
        self.resize(500, 600)
        
        layout = QVBoxLayout()
        
        tabs = QTabWidget()
        
        # General Settings Tab
        general_tab = QWidget()
        general_layout = QVBoxLayout()
        
        # Auto-Save Group
        autosave_group = QGroupBox("Auto-Save Configuration")
        autosave_form = QFormLayout()
        
        self.autosave_cb = QCheckBox("Enable Auto-Save")
        self.autosave_cb.setChecked(auto_save_enabled)
        autosave_form.addRow("Status:", self.autosave_cb)
        
        self.interval_spin = QSpinBox()
        self.interval_spin.setRange(1, 60)
        self.interval_spin.setSuffix(" minutes")
        self.interval_spin.setValue(auto_save_interval)
        self.interval_spin.setEnabled(auto_save_enabled)
        autosave_form.addRow("Save Interval:", self.interval_spin)
        
        self.keep_versions_spin = QSpinBox()
        self.keep_versions_spin.setRange(1, 100)
        self.keep_versions_spin.setSuffix(" versions")
        self.keep_versions_spin.setValue(keep_versions)
        self.keep_versions_spin.setEnabled(auto_save_enabled)
        autosave_form.addRow("Keep Versions:", self.keep_versions_spin)
        
        self.autosave_cb.toggled.connect(self.interval_spin.setEnabled)
        self.autosave_cb.toggled.connect(self.keep_versions_spin.setEnabled)
        
        autosave_group.setLayout(autosave_form)
        general_layout.addWidget(autosave_group)
        
        autosave_info = QLabel(
            "Auto-save periodically saves your work and maintains version history.\n\n"
            "• Only works for files that have been saved at least once\n"
            "• Version history is stored in ~/.excel_editor_versions\n"
            "• Older versions are automatically cleaned up"
        )
        autosave_info.setWordWrap(True)
        autosave_info.setStyleSheet("color: #888; font-style: italic; padding: 10px;")
        general_layout.addWidget(autosave_info)
        
        general_layout.addStretch()
        general_tab.setLayout(general_layout)
        tabs.addTab(general_tab, "General")
        
        # Theme Tab
        theme_tab = QWidget()
        theme_layout = QVBoxLayout()
        
        info_label = QLabel("Customize the application theme colors:")
        info_label.setStyleSheet("font-weight: bold; margin-bottom: 10px;")
        theme_layout.addWidget(info_label)
        
        bg_group = QGroupBox("Background Color")
        bg_layout = QHBoxLayout()
        self.bg_color = QColor(current_bg)
        self.bg_preview = QPushButton()
        self.bg_preview.setFixedSize(50, 50)
        self.bg_preview.setStyleSheet(f"background-color: {current_bg}; border: 2px solid #888;")
        self.bg_preview.clicked.connect(self.choose_bg_color)
        bg_layout.addWidget(self.bg_preview)
        self.bg_label = QLabel(f"Current: {current_bg}")
        bg_layout.addWidget(self.bg_label)
        bg_layout.addStretch()
        bg_change_btn = QPushButton("Change Color")
        bg_change_btn.clicked.connect(self.choose_bg_color)
        bg_layout.addWidget(bg_change_btn)
        bg_group.setLayout(bg_layout)
        theme_layout.addWidget(bg_group)
        
        accent_group = QGroupBox("Accent Color")
        accent_layout = QHBoxLayout()
        self.accent_color = QColor(current_accent)
        self.accent_preview = QPushButton()
        self.accent_preview.setFixedSize(50, 50)
        self.accent_preview.setStyleSheet(f"background-color: {current_accent}; border: 2px solid #888;")
        self.accent_preview.clicked.connect(self.choose_accent_color)
        accent_layout.addWidget(self.accent_preview)
        self.accent_label = QLabel(f"Current: {current_accent}")
        accent_layout.addWidget(self.accent_label)
        accent_layout.addStretch()
        accent_change_btn = QPushButton("Change Color")
        accent_change_btn.clicked.connect(self.choose_accent_color)
        accent_layout.addWidget(accent_change_btn)
        accent_group.setLayout(accent_layout)
        theme_layout.addWidget(accent_group)
        
        # Preset themes
        preset_group = QGroupBox("Preset Themes")
        preset_layout = QVBoxLayout()
        presets = [
            ("Default (Dark Coffee)", "#1A1916", "#404040"),
            ("Dark Blue & Cyan", "#1e2832", "#00d4ff"),
            ("Purple & Pink", "#2d1b3d", "#ff69b4"),
            ("Dark Gray & Red", "#2b2b2b", "#ff4444"),
            ("Navy & Gold", "#1a2332", "#ffd700"),
            ("Forest & Lime", "#1a3d2e", "#00ff00"),
            ("Green & Orange", "#313B2F", "#FBA002"),
            ("Paynes & Pearl", "#04111B", "#F5F112")
        ]
        for name, bg, accent in presets:
            preset_btn = QPushButton(name)
            preset_btn.clicked.connect(lambda checked, b=bg, a=accent: self.apply_preset(b, a))
            preset_layout.addWidget(preset_btn)
        preset_group.setLayout(preset_layout)
        theme_layout.addWidget(preset_group)
        
        theme_layout.addStretch()
        theme_tab.setLayout(theme_layout)
        tabs.addTab(theme_tab, "Theme")
        
        layout.addWidget(tabs)
        
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)
        
        self.setLayout(layout)
    
    def choose_bg_color(self):
        color = QColorDialog.getColor(self.bg_color, self, "Choose Background Color")
        if color.isValid():
            self.bg_color = color
            hex_color = color.name()
            self.bg_preview.setStyleSheet(f"background-color: {hex_color}; border: 2px solid #888;")
            self.bg_label.setText(f"Current: {hex_color}")
            
    def choose_accent_color(self):
        color = QColorDialog.getColor(self.accent_color, self, "Choose Accent Color")
        if color.isValid():
            self.accent_color = color
            hex_color = color.name()
            self.accent_preview.setStyleSheet(f"background-color: {hex_color}; border: 2px solid #888;")
            self.accent_label.setText(f"Current: {hex_color}")
            
    def apply_preset(self, bg, accent):
        self.bg_color = QColor(bg)
        self.accent_color = QColor(accent)
        self.bg_preview.setStyleSheet(f"background-color: {bg}; border: 2px solid #888;")
        self.bg_label.setText(f"Current: {bg}")
        self.accent_preview.setStyleSheet(f"background-color: {accent}; border: 2px solid #888;")
        self.accent_label.setText(f"Current: {accent}")
            
    def get_settings(self):
        return {
            'bg_color': self.bg_color.name(),
            'accent_color': self.accent_color.name(),
            'auto_save_enabled': self.autosave_cb.isChecked(),
            'auto_save_interval': self.interval_spin.value(),
            'keep_versions': self.keep_versions_spin.value()
        }