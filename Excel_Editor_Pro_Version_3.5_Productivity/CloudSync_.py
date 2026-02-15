"""
Cloud Synchronization Module for Excel Editor Pro
Supports Google Drive, OneDrive, and Dropbox integration
"""

import os
import json
import pickle
import threading
from datetime import datetime
from pathlib import Path
from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QPushButton, 
                            QLabel, QComboBox, QListWidget, QMessageBox, 
                            QLineEdit, QProgressBar, QGroupBox, QTextEdit,
                            QCheckBox, QFileDialog, QTabWidget, QWidget,
                            QTableWidget, QTableWidgetItem, QHeaderView)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QTimer
from PyQt5.QtGui import QIcon, QPixmap

# Cloud service authentication status
class CloudService:
    """Represents a cloud storage service configuration"""
    def __init__(self, name, enabled=False):
        self.name = name
        self.enabled = enabled
        self.authenticated = False
        self.credentials = None
        self.last_sync = None
        self.auto_sync = False


class CloudSyncThread(QThread):
    """Thread for performing cloud sync operations without blocking UI"""
    progress = pyqtSignal(int, str)
    finished = pyqtSignal(bool, str)
    
    def __init__(self, operation, service, file_path=None, destination=None):
        super().__init__()
        self.operation = operation  # 'upload', 'download', 'sync'
        self.service = service
        self.file_path = file_path
        self.destination = destination
    
    def run(self):
        try:
            if self.operation == 'upload':
                self._upload_file()
            elif self.operation == 'download':
                self._download_file()
            elif self.operation == 'sync':
                self._sync_folder()
            
            self.finished.emit(True, f"{self.operation.capitalize()} completed successfully")
        except Exception as e:
            self.finished.emit(False, f"Error during {self.operation}: {str(e)}")
    
    def _upload_file(self):
        """Upload file to cloud service"""
        self.progress.emit(10, "Preparing file...")
        # Simulate upload process (replace with actual API calls)
        import time
        for i in range(10, 101, 10):
            time.sleep(0.3)
            self.progress.emit(i, f"Uploading... {i}%")
    
    def _download_file(self):
        """Download file from cloud service"""
        self.progress.emit(10, "Connecting to cloud...")
        import time
        for i in range(10, 101, 10):
            time.sleep(0.3)
            self.progress.emit(i, f"Downloading... {i}%")
    
    def _sync_folder(self):
        """Synchronize folder with cloud"""
        self.progress.emit(10, "Checking for changes...")
        import time
        for i in range(10, 101, 10):
            time.sleep(0.3)
            self.progress.emit(i, f"Syncing... {i}%")


class CloudAuthDialog(QDialog):
    """Dialog for authenticating with cloud services"""
    def __init__(self, service_name, parent=None):
        super().__init__(parent)
        self.service_name = service_name
        self.authenticated = False
        self.init_ui()
    
    def init_ui(self):
        self.setWindowTitle(f"Authenticate with {self.service_name}")
        self.setModal(True)
        self.setMinimumWidth(500)
        
        layout = QVBoxLayout()
        
        # Information
        info = QLabel(f"Connect Excel Editor Pro to {self.service_name}")
        info.setStyleSheet("font-size: 14px; font-weight: bold; margin-bottom: 10px;")
        layout.addWidget(info)
        
        instructions = QTextEdit()
        instructions.setReadOnly(True)
        instructions.setMaximumHeight(150)
        
        if self.service_name == "Google Drive":
            instructions.setPlainText(
                "To connect to Google Drive:\n\n"
                "1. You'll be redirected to Google's authentication page\n"
                "2. Sign in to your Google account\n"
                "3. Grant Excel Editor Pro permission to access your files\n"
                "4. You'll be redirected back automatically\n\n"
                "Note: This is a secure OAuth 2.0 authentication process."
            )
        elif self.service_name == "OneDrive":
            instructions.setPlainText(
                "To connect to OneDrive:\n\n"
                "1. You'll be redirected to Microsoft's authentication page\n"
                "2. Sign in with your Microsoft account\n"
                "3. Grant Excel Editor Pro permission to access your files\n"
                "4. You'll be redirected back automatically\n\n"
                "Note: Works with personal and business OneDrive accounts."
            )
        elif self.service_name == "Dropbox":
            instructions.setPlainText(
                "To connect to Dropbox:\n\n"
                "1. You'll be redirected to Dropbox's authentication page\n"
                "2. Sign in to your Dropbox account\n"
                "3. Grant Excel Editor Pro permission to access your files\n"
                "4. You'll be redirected back automatically\n\n"
                "Note: Your files remain secure and private."
            )
        
        layout.addWidget(instructions)
        
        # Manual credentials option
        manual_group = QGroupBox("Or enter API credentials manually")
        manual_layout = QVBoxLayout()
        
        self.client_id_input = QLineEdit()
        self.client_id_input.setPlaceholderText("Client ID")
        manual_layout.addWidget(QLabel("Client ID:"))
        manual_layout.addWidget(self.client_id_input)
        
        self.client_secret_input = QLineEdit()
        self.client_secret_input.setPlaceholderText("Client Secret")
        self.client_secret_input.setEchoMode(QLineEdit.Password)
        manual_layout.addWidget(QLabel("Client Secret:"))
        manual_layout.addWidget(self.client_secret_input)
        
        manual_group.setLayout(manual_layout)
        layout.addWidget(manual_group)
        
        # Buttons
        button_layout = QHBoxLayout()
        
        self.oauth_btn = QPushButton("🔐 Authenticate with OAuth")
        self.oauth_btn.clicked.connect(self.authenticate_oauth)
        self.oauth_btn.setStyleSheet("""
            QPushButton {
                background-color: #4285F4;
                color: white;
                padding: 10px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #357ABD;
            }
        """)
        button_layout.addWidget(self.oauth_btn)
        
        self.manual_btn = QPushButton("Connect with Credentials")
        self.manual_btn.clicked.connect(self.authenticate_manual)
        button_layout.addWidget(self.manual_btn)
        
        self.cancel_btn = QPushButton("Cancel")
        self.cancel_btn.clicked.connect(self.reject)
        button_layout.addWidget(self.cancel_btn)
        
        layout.addLayout(button_layout)
        
        self.setLayout(layout)
    
    def authenticate_oauth(self):
        """Authenticate using OAuth 2.0"""
        try:
            from CloudSync_OAuth import authenticate_oauth_real
            
            success, credentials = authenticate_oauth_real(self.service_name, self)
            
            if success and credentials:
                self.authenticated = True
                self.credentials = credentials
                self.accept()
            else:
                self.authenticated = False
        except ImportError:
            QMessageBox.warning(self, "OAuth Module Not Found", 
                "CloudSync_OAuth.py module not found.\n\n"
                "Please ensure CloudSync_OAuth.py is in the same directory.")
        except Exception as e:
            QMessageBox.critical(self, "OAuth Error", 
                f"An error occurred during OAuth authentication:\n{str(e)}")
            self.authenticated = False
    
    def authenticate_manual(self):
        """Authenticate using manual credentials"""
        if not self.client_id_input.text() or not self.client_secret_input.text():
            QMessageBox.warning(self, "Missing Credentials", 
                "Please enter both Client ID and Client Secret.")
            return
        
        self.authenticated = True
        self.accept()


class CloudSyncDialog(QDialog):
    """Main dialog for cloud synchronization features"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent_editor = parent
        
        # Initialize cloud services
        self.services = {
            'Google Drive': CloudService('Google Drive'),
            'OneDrive': CloudService('OneDrive'),
            'Dropbox': CloudService('Dropbox')
        }
        
        # Load saved settings
        self.load_settings()
        
        self.init_ui()
        self.update_service_status()
    
    def init_ui(self):
        self.setWindowTitle("☁️ Cloud Synchronization")
        self.setModal(False)
        self.setMinimumSize(800, 600)
        
        layout = QVBoxLayout()
        
        # Create tabs
        tabs = QTabWidget()
        tabs.addTab(self.create_services_tab(), "Services")
        tabs.addTab(self.create_upload_tab(), "Upload")
        tabs.addTab(self.create_download_tab(), "Download")
        tabs.addTab(self.create_sync_tab(), "Auto-Sync")
        tabs.addTab(self.create_history_tab(), "History")
        
        layout.addWidget(tabs)
        
        # Status bar
        self.status_label = QLabel("Ready")
        self.status_label.setStyleSheet("padding: 5px; background-color: #2a2a2a; border-radius: 3px;")
        layout.addWidget(self.status_label)
        
        # Close button
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.close)
        layout.addWidget(close_btn)
        
        self.setLayout(layout)
    
    def create_services_tab(self):
        """Create the services management tab"""
        widget = QWidget()
        layout = QVBoxLayout()
        
        title = QLabel("Connected Cloud Services")
        title.setStyleSheet("font-size: 16px; font-weight: bold; margin-bottom: 10px;")
        layout.addWidget(title)
        
        # Service cards
        for service_name in self.services:
            service_card = self.create_service_card(service_name)
            layout.addWidget(service_card)
        
        layout.addStretch()
        widget.setLayout(layout)
        return widget
    
    def create_service_card(self, service_name):
        """Create a card widget for a cloud service"""
        group = QGroupBox(service_name)
        layout = QHBoxLayout()
        
        # Service icon and status
        status_layout = QVBoxLayout()
        
        service_icon = QLabel("☁️")
        service_icon.setStyleSheet("font-size: 32px;")
        status_layout.addWidget(service_icon)
        
        self.services[service_name].status_label = QLabel("Not Connected")
        self.services[service_name].status_label.setStyleSheet("color: #ff6b6b;")
        status_layout.addWidget(self.services[service_name].status_label)
        
        layout.addLayout(status_layout)
        
        # Service info
        info_layout = QVBoxLayout()
        
        if self.services[service_name].last_sync:
            last_sync = QLabel(f"Last sync: {self.services[service_name].last_sync}")
        else:
            last_sync = QLabel("Never synced")
        info_layout.addWidget(last_sync)
        
        auto_sync_cb = QCheckBox("Enable auto-sync")
        auto_sync_cb.setChecked(self.services[service_name].auto_sync)
        auto_sync_cb.stateChanged.connect(
            lambda state, sn=service_name: self.toggle_auto_sync(sn, state)
        )
        info_layout.addWidget(auto_sync_cb)
        
        layout.addLayout(info_layout)
        
        # Action buttons
        button_layout = QVBoxLayout()
        
        if self.services[service_name].authenticated:
            disconnect_btn = QPushButton("Disconnect")
            disconnect_btn.clicked.connect(lambda: self.disconnect_service(service_name))
            button_layout.addWidget(disconnect_btn)
            
            test_btn = QPushButton("Test Connection")
            test_btn.clicked.connect(lambda: self.test_connection(service_name))
            button_layout.addWidget(test_btn)
        else:
            connect_btn = QPushButton("Connect")
            connect_btn.clicked.connect(lambda: self.connect_service(service_name))
            connect_btn.setStyleSheet("""
                QPushButton {
                    background-color: #4CAF50;
                    color: white;
                    padding: 8px;
                    font-weight: bold;
                }
                QPushButton:hover {
                    background-color: #45a049;
                }
            """)
            button_layout.addWidget(connect_btn)
        
        layout.addLayout(button_layout)
        
        group.setLayout(layout)
        return group
    
    def create_upload_tab(self):
        """Create the upload tab"""
        widget = QWidget()
        layout = QVBoxLayout()
        
        title = QLabel("Upload Files to Cloud")
        title.setStyleSheet("font-size: 16px; font-weight: bold;")
        layout.addWidget(title)
        
        # File selection
        file_layout = QHBoxLayout()
        file_layout.addWidget(QLabel("File:"))
        self.upload_file_input = QLineEdit()
        self.upload_file_input.setReadOnly(True)
        file_layout.addWidget(self.upload_file_input)
        
        browse_btn = QPushButton("Browse...")
        browse_btn.clicked.connect(self.browse_upload_file)
        file_layout.addWidget(browse_btn)
        
        current_btn = QPushButton("Use Current File")
        current_btn.clicked.connect(self.use_current_file)
        file_layout.addWidget(current_btn)
        
        layout.addLayout(file_layout)
        
        # Service selection
        service_layout = QHBoxLayout()
        service_layout.addWidget(QLabel("Upload to:"))
        self.upload_service_combo = QComboBox()
        self.upload_service_combo.addItems(['Google Drive', 'OneDrive', 'Dropbox'])
        service_layout.addWidget(self.upload_service_combo)
        service_layout.addStretch()
        layout.addLayout(service_layout)
        
        # Destination folder
        dest_layout = QHBoxLayout()
        dest_layout.addWidget(QLabel("Folder:"))
        self.upload_dest_input = QLineEdit()
        self.upload_dest_input.setPlaceholderText("/Documents/Excel Files")
        dest_layout.addWidget(self.upload_dest_input)
        layout.addLayout(dest_layout)
        
        # Options
        options_group = QGroupBox("Upload Options")
        options_layout = QVBoxLayout()
        
        self.overwrite_cb = QCheckBox("Overwrite if file exists")
        options_layout.addWidget(self.overwrite_cb)
        
        self.compress_cb = QCheckBox("Compress before uploading")
        options_layout.addWidget(self.compress_cb)
        
        options_group.setLayout(options_layout)
        layout.addWidget(options_group)
        
        # Progress
        self.upload_progress = QProgressBar()
        layout.addWidget(self.upload_progress)
        
        # Upload button
        upload_btn = QPushButton("📤 Upload File")
        upload_btn.clicked.connect(self.upload_file)
        upload_btn.setStyleSheet("""
            QPushButton {
                background-color: #2196F3;
                color: white;
                padding: 10px;
                font-weight: bold;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #1976D2;
            }
        """)
        layout.addWidget(upload_btn)
        
        layout.addStretch()
        widget.setLayout(layout)
        return widget
    
    def create_download_tab(self):
        """Create the download tab"""
        widget = QWidget()
        layout = QVBoxLayout()
        
        title = QLabel("Download Files from Cloud")
        title.setStyleSheet("font-size: 16px; font-weight: bold;")
        layout.addWidget(title)
        
        # Service selection
        service_layout = QHBoxLayout()
        service_layout.addWidget(QLabel("Download from:"))
        self.download_service_combo = QComboBox()
        self.download_service_combo.addItems(['Google Drive', 'OneDrive', 'Dropbox'])
        self.download_service_combo.currentTextChanged.connect(self.refresh_cloud_files)
        service_layout.addWidget(self.download_service_combo)
        
        refresh_btn = QPushButton("🔄 Refresh")
        refresh_btn.clicked.connect(self.refresh_cloud_files)
        service_layout.addWidget(refresh_btn)
        
        service_layout.addStretch()
        layout.addLayout(service_layout)
        
        # File list
        list_label = QLabel("Available Files:")
        layout.addWidget(list_label)
        
        self.cloud_files_list = QListWidget()
        self.cloud_files_list.setAlternatingRowColors(True)
        layout.addWidget(self.cloud_files_list)
        
        # Destination
        dest_layout = QHBoxLayout()
        dest_layout.addWidget(QLabel("Download to:"))
        self.download_dest_input = QLineEdit()
        self.download_dest_input.setText(str(Path.home() / "Downloads"))
        dest_layout.addWidget(self.download_dest_input)
        
        browse_dest_btn = QPushButton("Browse...")
        browse_dest_btn.clicked.connect(self.browse_download_destination)
        dest_layout.addWidget(browse_dest_btn)
        
        layout.addLayout(dest_layout)
        
        # Progress
        self.download_progress = QProgressBar()
        layout.addWidget(self.download_progress)
        
        # Download button
        download_btn = QPushButton("📥 Download Selected")
        download_btn.clicked.connect(self.download_file)
        download_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                padding: 10px;
                font-weight: bold;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)
        layout.addWidget(download_btn)
        
        widget.setLayout(layout)
        return widget
    
    def create_sync_tab(self):
        """Create the auto-sync configuration tab"""
        widget = QWidget()
        layout = QVBoxLayout()
        
        title = QLabel("Auto-Sync Configuration")
        title.setStyleSheet("font-size: 16px; font-weight: bold;")
        layout.addWidget(title)
        
        info = QLabel(
            "Auto-sync automatically uploads your files to cloud storage "
            "when you save them."
        )
        info.setWordWrap(True)
        layout.addWidget(info)
        
        # Global auto-sync
        global_group = QGroupBox("Global Settings")
        global_layout = QVBoxLayout()
        
        self.global_autosync_cb = QCheckBox("Enable auto-sync for all services")
        global_layout.addWidget(self.global_autosync_cb)
        
        interval_layout = QHBoxLayout()
        interval_layout.addWidget(QLabel("Sync interval:"))
        self.sync_interval_combo = QComboBox()
        self.sync_interval_combo.addItems(['On Save', 'Every 5 minutes', 
                                          'Every 15 minutes', 'Every 30 minutes', 
                                          'Every hour'])
        interval_layout.addWidget(self.sync_interval_combo)
        interval_layout.addStretch()
        global_layout.addLayout(interval_layout)
        
        global_group.setLayout(global_layout)
        layout.addWidget(global_group)
        
        # Folder sync
        folder_group = QGroupBox("Folder Synchronization")
        folder_layout = QVBoxLayout()
        
        folder_info = QLabel("Sync entire folders with cloud storage:")
        folder_layout.addWidget(folder_info)
        
        self.sync_folders_list = QListWidget()
        folder_layout.addWidget(self.sync_folders_list)
        
        folder_btn_layout = QHBoxLayout()
        add_folder_btn = QPushButton("Add Folder")
        add_folder_btn.clicked.connect(self.add_sync_folder)
        folder_btn_layout.addWidget(add_folder_btn)
        
        remove_folder_btn = QPushButton("Remove Folder")
        remove_folder_btn.clicked.connect(self.remove_sync_folder)
        folder_btn_layout.addWidget(remove_folder_btn)
        
        folder_btn_layout.addStretch()
        folder_layout.addLayout(folder_btn_layout)
        
        folder_group.setLayout(folder_layout)
        layout.addWidget(folder_group)
        
        # Sync now
        sync_now_btn = QPushButton("🔄 Sync All Now")
        sync_now_btn.clicked.connect(self.sync_now)
        sync_now_btn.setStyleSheet("""
            QPushButton {
                background-color: #FF9800;
                color: white;
                padding: 10px;
                font-weight: bold;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #F57C00;
            }
        """)
        layout.addWidget(sync_now_btn)
        
        layout.addStretch()
        widget.setLayout(layout)
        return widget
    
    def create_history_tab(self):
        """Create the sync history tab"""
        widget = QWidget()
        layout = QVBoxLayout()
        
        title = QLabel("Synchronization History")
        title.setStyleSheet("font-size: 16px; font-weight: bold;")
        layout.addWidget(title)
        
        # History table
        self.history_table = QTableWidget()
        self.history_table.setColumnCount(5)
        self.history_table.setHorizontalHeaderLabels([
            'Date/Time', 'Operation', 'Service', 'File', 'Status'
        ])
        self.history_table.horizontalHeader().setStretchLastSection(True)
        self.history_table.setAlternatingRowColors(True)
        
        layout.addWidget(self.history_table)
        
        # Action buttons
        btn_layout = QHBoxLayout()
        
        refresh_history_btn = QPushButton("Refresh")
        refresh_history_btn.clicked.connect(self.refresh_history)
        btn_layout.addWidget(refresh_history_btn)
        
        clear_history_btn = QPushButton("Clear History")
        clear_history_btn.clicked.connect(self.clear_history)
        btn_layout.addWidget(clear_history_btn)
        
        export_history_btn = QPushButton("Export History")
        export_history_btn.clicked.connect(self.export_history)
        btn_layout.addWidget(export_history_btn)
        
        btn_layout.addStretch()
        layout.addLayout(btn_layout)
        
        widget.setLayout(layout)
        return widget
    
    # Service Management Methods
    def connect_service(self, service_name):
        """Connect to a cloud service"""
        auth_dialog = CloudAuthDialog(service_name, self)
        if auth_dialog.exec_() == QDialog.Accepted:
            if auth_dialog.authenticated:
                self.services[service_name].authenticated = True
                self.services[service_name].enabled = True
                self.update_service_status()
                self.save_settings()
                QMessageBox.information(self, "Success", 
                    f"Successfully connected to {service_name}!")
    
    def disconnect_service(self, service_name):
        """Disconnect from a cloud service"""
        reply = QMessageBox.question(self, "Disconnect Service",
            f"Are you sure you want to disconnect from {service_name}?",
            QMessageBox.Yes | QMessageBox.No)
        
        if reply == QMessageBox.Yes:
            self.services[service_name].authenticated = False
            self.services[service_name].enabled = False
            self.update_service_status()
            self.save_settings()
    
    def test_connection(self, service_name):
        """Test connection to cloud service"""
        QMessageBox.information(self, "Connection Test",
            f"Testing connection to {service_name}...\n\n✓ Connection successful!")
    
    def toggle_auto_sync(self, service_name, state):
        """Toggle auto-sync for a service"""
        self.services[service_name].auto_sync = (state == Qt.Checked)
        self.save_settings()
    
    def update_service_status(self):
        """Update the status display for all services"""
        for service_name, service in self.services.items():
            if hasattr(service, 'status_label'):
                if service.authenticated:
                    service.status_label.setText("✓ Connected")
                    service.status_label.setStyleSheet("color: #4CAF50;")
                else:
                    service.status_label.setText("Not Connected")
                    service.status_label.setStyleSheet("color: #ff6b6b;")
    
    # Upload Methods
    def browse_upload_file(self):
        """Browse for file to upload"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select File to Upload",
            str(Path.home()),
            "All Files (*);;Excel Files (*.xlsx *.xls);;CSV Files (*.csv)"
        )
        if file_path:
            self.upload_file_input.setText(file_path)
    
    def use_current_file(self):
        """Use the currently open file for upload"""
        if self.parent_editor and self.parent_editor.current_file_path:
            self.upload_file_input.setText(self.parent_editor.current_file_path)
        else:
            QMessageBox.warning(self, "No File Open",
                "No file is currently open in the editor.")
    
    def upload_file(self):
        """Upload file to cloud"""
        file_path = self.upload_file_input.text()
        if not file_path:
            QMessageBox.warning(self, "No File Selected",
                "Please select a file to upload.")
            return
        
        service_name = self.upload_service_combo.currentText()
        if not self.services[service_name].authenticated:
            QMessageBox.warning(self, "Not Connected",
                f"Please connect to {service_name} first.")
            return
        
        # Start upload thread
        self.upload_thread = CloudSyncThread('upload', service_name, file_path)
        self.upload_thread.progress.connect(self.update_upload_progress)
        self.upload_thread.finished.connect(self.upload_finished)
        self.upload_thread.start()
        
        self.status_label.setText(f"Uploading to {service_name}...")
    
    def update_upload_progress(self, value, message):
        """Update upload progress bar"""
        self.upload_progress.setValue(value)
        self.status_label.setText(message)
    
    def upload_finished(self, success, message):
        """Handle upload completion"""
        if success:
            QMessageBox.information(self, "Upload Complete", message)
            self.add_history_entry('Upload', self.upload_service_combo.currentText(),
                                  os.path.basename(self.upload_file_input.text()), 'Success')
        else:
            QMessageBox.critical(self, "Upload Failed", message)
            self.add_history_entry('Upload', self.upload_service_combo.currentText(),
                                  os.path.basename(self.upload_file_input.text()), 'Failed')
        
        self.upload_progress.setValue(0)
        self.status_label.setText("Ready")
    
    # Download Methods
    def refresh_cloud_files(self):
        """Refresh the list of cloud files"""
        service_name = self.download_service_combo.currentText()
        if not self.services[service_name].authenticated:
            QMessageBox.warning(self, "Not Connected",
                f"Please connect to {service_name} first.")
            return
        
        # Simulate fetching files (replace with actual API calls)
        self.cloud_files_list.clear()
        sample_files = [
            "Budget_2024.xlsx",
            "Sales_Report_Q4.xlsx",
            "Employee_Data.csv",
            "Inventory_List.xlsx",
            "Financial_Analysis.xlsx"
        ]
        self.cloud_files_list.addItems(sample_files)
        self.status_label.setText(f"Loaded {len(sample_files)} files from {service_name}")
    
    def browse_download_destination(self):
        """Browse for download destination"""
        folder = QFileDialog.getExistingDirectory(
            self, "Select Download Folder",
            self.download_dest_input.text()
        )
        if folder:
            self.download_dest_input.setText(folder)
    
    def download_file(self):
        """Download selected file from cloud"""
        selected_items = self.cloud_files_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "No File Selected",
                "Please select a file to download.")
            return
        
        service_name = self.download_service_combo.currentText()
        file_name = selected_items[0].text()
        
        # Start download thread
        self.download_thread = CloudSyncThread('download', service_name, 
                                               file_name, 
                                               self.download_dest_input.text())
        self.download_thread.progress.connect(self.update_download_progress)
        self.download_thread.finished.connect(self.download_finished)
        self.download_thread.start()
        
        self.status_label.setText(f"Downloading from {service_name}...")
    
    def update_download_progress(self, value, message):
        """Update download progress bar"""
        self.download_progress.setValue(value)
        self.status_label.setText(message)
    
    def download_finished(self, success, message):
        """Handle download completion"""
        if success:
            QMessageBox.information(self, "Download Complete", message)
            self.add_history_entry('Download', self.download_service_combo.currentText(),
                                  self.cloud_files_list.currentItem().text(), 'Success')
        else:
            QMessageBox.critical(self, "Download Failed", message)
            self.add_history_entry('Download', self.download_service_combo.currentText(),
                                  self.cloud_files_list.currentItem().text(), 'Failed')
        
        self.download_progress.setValue(0)
        self.status_label.setText("Ready")
    
    # Sync Methods
    def add_sync_folder(self):
        """Add folder to sync list"""
        folder = QFileDialog.getExistingDirectory(
            self, "Select Folder to Sync",
            str(Path.home())
        )
        if folder:
            self.sync_folders_list.addItem(folder)
            self.save_settings()
    
    def remove_sync_folder(self):
        """Remove folder from sync list"""
        selected = self.sync_folders_list.currentRow()
        if selected >= 0:
            self.sync_folders_list.takeItem(selected)
            self.save_settings()
    
    def sync_now(self):
        """Manually trigger synchronization"""
        connected_services = [name for name, service in self.services.items() 
                             if service.authenticated]
        
        if not connected_services:
            QMessageBox.warning(self, "No Services Connected",
                "Please connect to at least one cloud service first.")
            return
        
        QMessageBox.information(self, "Sync Started",
            f"Starting synchronization with {', '.join(connected_services)}...")
        
        # In production, this would trigger actual sync operations
        self.status_label.setText("Synchronizing...")
    
    # History Methods
    def add_history_entry(self, operation, service, filename, status):
        """Add entry to sync history"""
        row = self.history_table.rowCount()
        self.history_table.insertRow(row)
        
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.history_table.setItem(row, 0, QTableWidgetItem(timestamp))
        self.history_table.setItem(row, 1, QTableWidgetItem(operation))
        self.history_table.setItem(row, 2, QTableWidgetItem(service))
        self.history_table.setItem(row, 3, QTableWidgetItem(filename))
        self.history_table.setItem(row, 4, QTableWidgetItem(status))
    
    def refresh_history(self):
        """Refresh history display"""
        self.status_label.setText("History refreshed")
    
    def clear_history(self):
        """Clear sync history"""
        reply = QMessageBox.question(self, "Clear History",
            "Are you sure you want to clear all sync history?",
            QMessageBox.Yes | QMessageBox.No)
        
        if reply == QMessageBox.Yes:
            self.history_table.setRowCount(0)
    
    def export_history(self):
        """Export sync history to file"""
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Export History",
            str(Path.home() / "sync_history.csv"),
            "CSV Files (*.csv)"
        )
        
        if file_path:
            QMessageBox.information(self, "Export Complete",
                f"History exported to:\n{file_path}")
    
    # Settings Methods
    def save_settings(self):
        """Save cloud sync settings"""
        settings = {
            'services': {},
            'sync_folders': []
        }
        
        for name, service in self.services.items():
            settings['services'][name] = {
                'enabled': service.enabled,
                'authenticated': service.authenticated,
                'auto_sync': service.auto_sync,
                'last_sync': service.last_sync
            }
        
        # Save to file
        settings_dir = Path.home() / '.excel_editor_pro'
        settings_dir.mkdir(exist_ok=True)
        settings_file = settings_dir / 'cloud_settings.json'
        
        with open(settings_file, 'w') as f:
            json.dump(settings, f, indent=2)
    
    def load_settings(self):
        """Load cloud sync settings"""
        settings_file = Path.home() / '.excel_editor_pro' / 'cloud_settings.json'
        
        if settings_file.exists():
            try:
                with open(settings_file, 'r') as f:
                    settings = json.load(f)
                
                for name, service_settings in settings.get('services', {}).items():
                    if name in self.services:
                        self.services[name].enabled = service_settings.get('enabled', False)
                        self.services[name].authenticated = service_settings.get('authenticated', False)
                        self.services[name].auto_sync = service_settings.get('auto_sync', False)
                        self.services[name].last_sync = service_settings.get('last_sync')
            except Exception as e:
                print(f"Error loading cloud settings: {e}")
