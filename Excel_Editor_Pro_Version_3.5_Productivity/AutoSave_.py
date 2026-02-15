"""
Auto-Save Manager with Version History
Provides continuous auto-save functionality with version history tracking
"""

import os
import shutil
from datetime import datetime
from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QPushButton, 
                             QListWidget, QLabel, QMessageBox, QGroupBox,
                             QSpinBox, QCheckBox)
from PyQt5.QtCore import QSettings, QTimer, Qt


class AutoSaveManager:
    """Manages auto-save functionality with version history"""
    
    def __init__(self, parent):
        self.parent = parent
        self.settings = QSettings("DataTools", "ExcelEditor")
        
        # Auto-save settings
        self.enabled = self.settings.value("auto_save/enabled", True, type=bool)
        self.interval = int(self.settings.value("auto_save/interval", 5))  # minutes
        self.keep_versions = int(self.settings.value("auto_save/keep_versions", 10))
        
        # Version history directory
        self.version_dir = os.path.join(os.path.expanduser("~"), ".excel_editor_versions")
        if not os.path.exists(self.version_dir):
            os.makedirs(self.version_dir)
        
        # Timer for auto-save
        self.timer = QTimer()
        self.timer.timeout.connect(self.perform_auto_save)
        if self.enabled:
            self.timer.start(self.interval * 60 * 1000)
    
    def perform_auto_save(self):
        """Perform auto-save operation"""
        try:
            if (self.parent.df is not None and 
                self.parent.is_modified and 
                self.parent.current_file_path):
                
                # Save current state
                self.save_version()
                
                # Save the actual file
                self.parent.save_file()
                
                # Show notification
                self.parent.statusBar().showMessage("Auto-saved", 2000)
                
        except Exception as e:
            print(f"Auto-save error: {e}")
    
    def save_version(self):
        """Save a version to history"""
        try:
            if not self.parent.current_file_path:
                return
            
            # Create version filename
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            base_name = os.path.basename(self.parent.current_file_path)
            name_without_ext = os.path.splitext(base_name)[0]
            ext = os.path.splitext(base_name)[1]
            
            version_filename = f"{name_without_ext}_v{timestamp}{ext}"
            version_path = os.path.join(self.version_dir, version_filename)
            
            # Save version
            include_index = self.parent.include_index_cb.isChecked()
            
            if self.parent.current_file_path.endswith('.csv'):
                self.parent.df.to_csv(version_path, index=include_index)
            else:
                self.parent.df.to_excel(version_path, index=include_index)
            
            # Clean old versions
            self.cleanup_old_versions(name_without_ext)
            
        except Exception as e:
            print(f"Error saving version: {e}")
    
    def cleanup_old_versions(self, base_name):
        """Remove old versions beyond the keep limit"""
        try:
            # Get all versions for this file
            versions = []
            for f in os.listdir(self.version_dir):
                if f.startswith(base_name + "_v"):
                    full_path = os.path.join(self.version_dir, f)
                    versions.append((os.path.getmtime(full_path), full_path))
            
            # Sort by modification time (oldest first)
            versions.sort()
            
            # Remove oldest if we exceed the limit
            while len(versions) > self.keep_versions:
                _, old_file = versions.pop(0)
                try:
                    os.remove(old_file)
                except:
                    pass
        except Exception as e:
            print(f"Error cleaning old versions: {e}")
    
    def get_versions(self, file_path):
        """Get list of versions for a file"""
        try:
            base_name = os.path.splitext(os.path.basename(file_path))[0]
            versions = []
            
            for f in os.listdir(self.version_dir):
                if f.startswith(base_name + "_v"):
                    full_path = os.path.join(self.version_dir, f)
                    mtime = os.path.getmtime(full_path)
                    size = os.path.getsize(full_path)
                    versions.append({
                        'filename': f,
                        'path': full_path,
                        'modified': datetime.fromtimestamp(mtime),
                        'size': size
                    })
            
            # Sort by modified time (newest first)
            versions.sort(key=lambda x: x['modified'], reverse=True)
            return versions
            
        except Exception as e:
            print(f"Error getting versions: {e}")
            return []
    
    def update_settings(self, enabled, interval, keep_versions):
        """Update auto-save settings"""
        self.enabled = enabled
        self.interval = interval
        self.keep_versions = keep_versions
        
        # Save to settings
        self.settings.setValue("auto_save/enabled", enabled)
        self.settings.setValue("auto_save/interval", interval)
        self.settings.setValue("auto_save/keep_versions", keep_versions)
        
        # Update timer
        if enabled:
            self.timer.start(interval * 60 * 1000)
        else:
            self.timer.stop()


class VersionHistoryDialog(QDialog):
    """Dialog to view and restore version history"""
    
    def __init__(self, parent, auto_save_manager):
        super().__init__(parent)
        self.parent = parent
        self.auto_save_manager = auto_save_manager
        self.init_ui()
        self.load_versions()
    
    def init_ui(self):
        self.setWindowTitle("Version History")
        self.setGeometry(200, 200, 600, 500)
        
        layout = QVBoxLayout()
        
        # Info label
        info_label = QLabel("Select a version to restore:")
        layout.addWidget(info_label)
        
        # Version list
        self.version_list = QListWidget()
        self.version_list.itemDoubleClicked.connect(self.restore_version)
        layout.addWidget(self.version_list)
        
        # Buttons
        button_layout = QHBoxLayout()
        
        restore_btn = QPushButton("Restore Selected")
        restore_btn.clicked.connect(self.restore_version)
        button_layout.addWidget(restore_btn)
        
        delete_btn = QPushButton("Delete Version")
        delete_btn.clicked.connect(self.delete_version)
        button_layout.addWidget(delete_btn)
        
        button_layout.addStretch()
        
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.close)
        button_layout.addWidget(close_btn)
        
        layout.addLayout(button_layout)
        
        self.setLayout(layout)
    
    def load_versions(self):
        """Load version history"""
        self.version_list.clear()
        
        if not self.parent.current_file_path:
            self.version_list.addItem("No file loaded")
            return
        
        versions = self.auto_save_manager.get_versions(self.parent.current_file_path)
        
        if not versions:
            self.version_list.addItem("No versions found")
            return
        
        for version in versions:
            # Format display
            time_str = version['modified'].strftime("%Y-%m-%d %H:%M:%S")
            size_kb = version['size'] / 1024
            display_text = f"{time_str} - {size_kb:.1f} KB"
            
            item = self.version_list.addItem(display_text)
            # Store version data
            self.version_list.item(self.version_list.count() - 1).setData(Qt.UserRole, version)
    
    def restore_version(self):
        """Restore selected version"""
        current_item = self.version_list.currentItem()
        if not current_item:
            return
        
        version = current_item.data(Qt.UserRole)
        if not version:
            return
        
        reply = QMessageBox.question(
            self,
            "Restore Version",
            f"Are you sure you want to restore this version?\n\n"
            f"Date: {version['modified'].strftime('%Y-%m-%d %H:%M:%S')}\n"
            f"Current unsaved changes will be lost.",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            try:
                # Load the version
                import pandas as pd
                if version['path'].endswith('.csv'):
                    df = pd.read_csv(version['path'])
                else:
                    df = pd.read_excel(version['path'])
                
                # Update parent
                self.parent.df = df
                self.parent.update_table()
                self.parent.is_modified = True
                
                QMessageBox.information(self, "Success", "Version restored successfully!")
                self.close()
                
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to restore version:\n{str(e)}")
    
    def delete_version(self):
        """Delete selected version"""
        current_item = self.version_list.currentItem()
        if not current_item:
            return
        
        version = current_item.data(Qt.UserRole)
        if not version:
            return
        
        reply = QMessageBox.question(
            self,
            "Delete Version",
            f"Are you sure you want to delete this version?\n\n"
            f"Date: {version['modified'].strftime('%Y-%m-%d %H:%M:%S')}\n"
            f"This action cannot be undone.",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            try:
                os.remove(version['path'])
                self.load_versions()
                QMessageBox.information(self, "Success", "Version deleted successfully!")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to delete version:\n{str(e)}")


class AutoSaveSettingsDialog(QDialog):
    """Dialog for configuring auto-save settings"""
    
    def __init__(self, parent, auto_save_manager):
        super().__init__(parent)
        self.auto_save_manager = auto_save_manager
        self.init_ui()
    
    def init_ui(self):
        self.setWindowTitle("Auto-Save Settings")
        self.setGeometry(300, 300, 400, 250)
        
        layout = QVBoxLayout()
        
        # Enable/disable auto-save
        self.enable_checkbox = QCheckBox("Enable Auto-Save")
        self.enable_checkbox.setChecked(self.auto_save_manager.enabled)
        layout.addWidget(self.enable_checkbox)
        
        # Interval setting
        interval_group = QGroupBox("Auto-Save Interval")
        interval_layout = QHBoxLayout()
        
        interval_layout.addWidget(QLabel("Save every:"))
        
        self.interval_spinbox = QSpinBox()
        self.interval_spinbox.setMinimum(1)
        self.interval_spinbox.setMaximum(60)
        self.interval_spinbox.setValue(self.auto_save_manager.interval)
        self.interval_spinbox.setSuffix(" minutes")
        interval_layout.addWidget(self.interval_spinbox)
        
        interval_group.setLayout(interval_layout)
        layout.addWidget(interval_group)
        
        # Version history setting
        version_group = QGroupBox("Version History")
        version_layout = QHBoxLayout()
        
        version_layout.addWidget(QLabel("Keep last:"))
        
        self.versions_spinbox = QSpinBox()
        self.versions_spinbox.setMinimum(1)
        self.versions_spinbox.setMaximum(100)
        self.versions_spinbox.setValue(self.auto_save_manager.keep_versions)
        self.versions_spinbox.setSuffix(" versions")
        version_layout.addWidget(self.versions_spinbox)
        
        version_group.setLayout(version_layout)
        layout.addWidget(version_group)
        
        layout.addStretch()
        
        # Buttons
        button_layout = QHBoxLayout()
        
        save_btn = QPushButton("Save")
        save_btn.clicked.connect(self.save_settings)
        button_layout.addWidget(save_btn)
        
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        button_layout.addWidget(cancel_btn)
        
        layout.addLayout(button_layout)
        
        self.setLayout(layout)
    
    def save_settings(self):
        """Save auto-save settings"""
        self.auto_save_manager.update_settings(
            self.enable_checkbox.isChecked(),
            self.interval_spinbox.value(),
            self.versions_spinbox.value()
        )
        self.accept()
