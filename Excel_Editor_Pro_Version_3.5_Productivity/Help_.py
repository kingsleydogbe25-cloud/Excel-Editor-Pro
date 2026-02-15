"""
Help Dialog: Comprehensive help system with multiple sections
"""
from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QTabWidget, 
                             QTextBrowser, QDialogButtonBox, QPushButton,
                             QLabel, QTableWidget, QTableWidgetItem, QHeaderView,
                             QWidget, QScrollArea)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont, QKeySequence


class HelpDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Excel Editor Pro - Help & Documentation")
        self.setModal(False)
        self.resize(900, 700)
        
        layout = QVBoxLayout()
        
        # Create tab widget
        self.tab_widget = QTabWidget()
        
        # Add tabs
        self.tab_widget.addTab(self.create_quick_start_tab(), "Quick Start")
        self.tab_widget.addTab(self.create_keyboard_shortcuts_tab(), "Keyboard Shortcuts")
        self.tab_widget.addTab(self.create_features_tab(), "Features Guide")
        self.tab_widget.addTab(self.create_cloud_sync_tab(), "☁️ Cloud Sync")
        self.tab_widget.addTab(self.create_ai_features_tab(), "🤖 AI Features")
        self.tab_widget.addTab(self.create_formatting_tab(), "Formatting Guide")
        self.tab_widget.addTab(self.create_transformation_tab(), "Data Transformation")
        self.tab_widget.addTab(self.create_troubleshooting_tab(), "Troubleshooting")
        self.tab_widget.addTab(self.create_about_tab(), "About")
        
        layout.addWidget(self.tab_widget)
        
        # Button box
        button_box = QDialogButtonBox(QDialogButtonBox.Close)
        button_box.rejected.connect(self.close)
        layout.addWidget(button_box)
        
        self.setLayout(layout)
    
    def create_quick_start_tab(self):
        """Quick start guide"""
        widget = QTextBrowser()
        widget.setOpenExternalLinks(True)
        
        content = """
        <h1>Quick Start Guide</h1>
        
        <h2>Getting Started</h2>
        <p>Welcome to Excel Editor Pro! This guide will help you get started quickly.</p>
        
        <h3>1. Creating a New File</h3>
        <ol>
            <li>Click the <b>New File</b> button in the toolbar</li>
            <li>Choose file type (Excel .xlsx or CSV .csv)</li>
            <li>Set initial dimensions (rows and columns)</li>
            <li>Choose whether to include headers</li>
            <li>Click OK to create your file</li>
        </ol>
        
        <h3>2. Opening an Existing File</h3>
        <ol>
            <li>Click the <b>Open File</b> button</li>
            <li>Browse to your Excel (.xlsx, .xls) or CSV file</li>
            <li>If the file has multiple sheets, select the one you want to edit</li>
            <li>The file will load into the editor</li>
        </ol>
        
        <h3>3. Editing Data</h3>
        <ul>
            <li><b>Edit cells:</b> Double-click any cell to edit its content</li>
            <li><b>Add rows:</b> Click "Add Row" button and choose position</li>
            <li><b>Add columns:</b> Click "Add Column" and specify details</li>
            <li><b>Delete rows:</b> Select a row and click "Delete Row"</li>
            <li><b>Reorder rows:</b> Use "Move Up ↑" and "Move Down ↓" buttons</li>
        </ul>
        
        <h3>4. Filtering Data</h3>
        <ol>
            <li>Select a column from the filter dropdown</li>
            <li>Enter search text in the filter box</li>
            <li>Click "Apply Filter" to show matching rows</li>
            <li>Click "Clear Filter" to show all data again</li>
        </ol>
        
        <h3>5. Formatting Columns</h3>
        <ol>
            <li>Click the <b>Format Columns</b> button</li>
            <li>Select a column to format</li>
            <li>Set fonts, colors, alignment, and number formats</li>
            <li>Click Apply - formatting appears immediately!</li>
        </ol>
        
        <h3>6. Saving Your Work</h3>
        <ul>
            <li><b>Save:</b> Click "Save" to save changes to the current file</li>
            <li><b>Save As:</b> Click "Save As" to save with a new name or format</li>
            <li><b>Auto-save:</b> Enable in Settings for automatic saves</li>
            <li><b>Cloud Sync:</b> Upload to Google Drive, OneDrive, or Dropbox</li>
        </ul>
        
        <h3>7. Cloud Synchronization (NEW!)</h3>
        <ol>
            <li>Click the <b>☁️ Cloud Sync</b> button or press <b>Ctrl+Shift+U</b></li>
            <li>Go to "Services" tab and click "Connect" on your cloud service</li>
            <li>Authenticate using OAuth (secure login)</li>
            <li>Use <b>Ctrl+U</b> to quickly upload current file</li>
            <li>Enable auto-sync for automatic backup on save</li>
        </ol>
        
        <h3>8. Exporting</h3>
        <ul>
            <li><b>Export to Excel:</b> Save As and choose .xlsx format</li>
            <li><b>Export to CSV:</b> Save As and choose .csv format</li>
            <li><b>Export to Word:</b> Use File → Export to DOC for reports</li>
        </ul>
        
        <h2>Pro Tips</h2>
        <ul>
            <li>💡 Use <b>Ctrl+S</b> to save quickly</li>
            <li>💡 Use <b>Ctrl+U</b> to quick upload to cloud</li>
            <li>💡 Enable auto-sync + auto-save for complete data protection</li>
            <li>💡 Large files? Only first 1000 rows display, but all data is preserved</li>
            <li>💡 Right-click the table for context menu options</li>
            <li>💡 Check the info panel on the right for dataset statistics</li>
            <li>💡 Use filters to work with specific data subsets</li>
            <li>💡 Connect to multiple cloud services for redundancy</li>
        </ul>
        """
        
        widget.setHtml(content)
        return widget
    
    def create_keyboard_shortcuts_tab(self):
        """Keyboard shortcuts reference"""
        widget = QWidget()
        layout = QVBoxLayout()
        
        # Title
        title = QLabel("<h2>Keyboard Shortcuts</h2>")
        layout.addWidget(title)
        
        # Create table for shortcuts
        table = QTableWidget()
        table.setColumnCount(3)
        table.setHorizontalHeaderLabels(["Action", "Shortcut", "Description"])
        
        shortcuts = [
            ("New File", "Ctrl+N", "Create a new Excel or CSV file"),
            ("Open File", "Ctrl+O", "Open an existing file"),
            ("Save", "Ctrl+S", "Save current file"),
            ("Save As", "Ctrl+Shift+S", "Save with new name"),
            ("Close", "Ctrl+W", "Close current file"),
            ("Quit", "Ctrl+Q", "Exit application"),
            ("", "", ""),
            ("Undo", "Ctrl+Z", "Undo last change (if available)"),
            ("Redo", "Ctrl+Y", "Redo last undone change"),
            ("", "", ""),
            ("Find/Filter", "Ctrl+F", "Focus on filter input"),
            ("Statistics", "Ctrl+I", "Show dataset statistics"),
            ("Settings", "Ctrl+,", "Open settings dialog"),
            ("", "", ""),
            ("Add Row", "Ctrl+R", "Add a new row"),
            ("Add Column", "Ctrl+Shift+C", "Add a new column"),
            ("Delete Row", "Delete", "Delete selected row"),
            ("", "", ""),
            ("Format Columns", "Ctrl+Shift+F", "Open column formatting"),
            ("Advanced Format", "Ctrl+Alt+F", "Open advanced formatting"),
            ("Transform Data", "Ctrl+T", "Open data transformation"),
            ("", "", ""),
            ("AI Assistant", "Ctrl+Shift+A", "Open AI features dialog"),
            ("Quick AI Analysis", "Ctrl+Alt+A", "Run quick AI analysis"),
            ("", "", ""),
            ("Cloud Sync", "Ctrl+Shift+U", "Open cloud synchronization"),
            ("Quick Upload", "Ctrl+U", "Quick upload to cloud"),
            ("", "", ""),
            ("Help", "F1", "Open this help dialog"),
            ("Refresh View", "F5", "Refresh the table display"),
        ]
        
        table.setRowCount(len(shortcuts))
        
        for row, (action, shortcut, description) in enumerate(shortcuts):
            table.setItem(row, 0, QTableWidgetItem(action))
            table.setItem(row, 1, QTableWidgetItem(shortcut))
            table.setItem(row, 2, QTableWidgetItem(description))
            
            # Make shortcut column bold
            if shortcut:
                font = QFont()
                font.setBold(True)
                table.item(row, 1).setFont(font)
        
        table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        table.setSelectionBehavior(QTableWidget.SelectRows)
        table.setAlternatingRowColors(True)
        
        layout.addWidget(table)
        
        # Print shortcut tip
        tip = QLabel("<i>💡 Tip: Print this page or take a screenshot for quick reference!</i>")
        layout.addWidget(tip)
        
        widget.setLayout(layout)
        return widget
    
    def create_features_tab(self):
        """Features guide"""
        widget = QTextBrowser()
        
        content = """
        <h1>Features Guide</h1>
        
        <h2>File Operations</h2>
        
        <h3>Creating Files</h3>
        <p>Create new Excel (.xlsx) or CSV (.csv) files with custom dimensions.</p>
        <ul>
            <li>Choose number of rows and columns</li>
            <li>Option to include/exclude headers</li>
            <li>Start with clean, empty data</li>
        </ul>
        
        <h3>Opening Files</h3>
        <p>Open and edit existing spreadsheet files.</p>
        <ul>
            <li>Supports Excel (.xlsx, .xls) and CSV (.csv) formats</li>
            <li>Multi-sheet Excel support - choose which sheet to edit</li>
            <li>Large file handling - displays first 1000 rows efficiently</li>
            <li>Preserves all data even when not all is displayed</li>
        </ul>
        
        <h3>Saving Files</h3>
        <ul>
            <li><b>Save:</b> Quick save to current file</li>
            <li><b>Save As:</b> Save with new name or convert formats</li>
            <li><b>Auto-save:</b> Configurable automatic saving (Settings)</li>
            <li><b>Format preservation:</b> Column formatting saved to Excel files</li>
        </ul>
        
        <h2>Data Editing</h2>
        
        <h3>Cell Editing</h3>
        <ul>
            <li>Double-click any cell to edit</li>
            <li>Automatic type detection (numbers, text)</li>
            <li>Support for empty cells</li>
        </ul>
        
        <h3>Row Operations</h3>
        <ul>
            <li><b>Add Row:</b> Insert at top, bottom, or before/after selection</li>
            <li><b>Delete Row:</b> Remove selected rows with confirmation</li>
            <li><b>Move Up/Down:</b> Reorder rows by moving them</li>
        </ul>
        
        <h3>Column Operations</h3>
        <ul>
            <li><b>Add Column:</b> Insert with custom name and position</li>
            <li><b>Format Column:</b> Apply formatting to entire columns</li>
            <li><b>Auto-resize:</b> Double-click column border to auto-fit</li>
        </ul>
        
        <h2>Filtering & Search</h2>
        <ul>
            <li>Filter by any column</li>
            <li>Case-insensitive text search</li>
            <li>Clear filter to restore full view</li>
            <li>Filtered view shows matching rows only</li>
        </ul>
        
        <h2>Statistics & Analysis</h2>
        <p>View comprehensive statistics about your data:</p>
        <ul>
            <li>Row and column counts</li>
            <li>Data types distribution</li>
            <li>Missing values detection</li>
            <li>Memory usage information</li>
            <li>Numeric column statistics (mean, median, std dev, etc.)</li>
        </ul>
        
        <h2>Formatting</h2>
        
        <h3>Basic Column Formatting</h3>
        <ul>
            <li>Font family, size, and style</li>
            <li>Text and background colors</li>
            <li>Text alignment (horizontal & vertical)</li>
            <li>Column width settings</li>
            <li>Header customization</li>
        </ul>
        
        <h3>Number Formats</h3>
        <ul>
            <li>General, Number, Currency</li>
            <li>Percentage, Date, Time</li>
            <li>Scientific notation</li>
            <li>Custom decimal places</li>
            <li>Thousands separators</li>
            <li>Currency symbols</li>
        </ul>
        
        <h3>Advanced Formatting</h3>
        <ul>
            <li>Apply to multiple cells at once</li>
            <li>Cell borders and styles</li>
            <li>Conditional formatting options</li>
            <li>Row height adjustments</li>
        </ul>
        
        <h2>Data Transformation</h2>
        <p>Powerful tools to clean and transform your data:</p>
        <ul>
            <li>Text transformations (uppercase, lowercase, title case)</li>
            <li>Number operations (multiply, add, round)</li>
            <li>Date formatting and manipulation</li>
            <li>Remove duplicates</li>
            <li>Fill missing values</li>
            <li>Split and merge columns</li>
            <li>Custom calculations</li>
        </ul>
        
        <h2>Export Options</h2>
        <ul>
            <li><b>Excel Export:</b> Full formatting preservation</li>
            <li><b>CSV Export:</b> Plain text data</li>
            <li><b>Word Document:</b> Create formatted reports with selected columns</li>
        </ul>
        
        <h2>Customization</h2>
        
        <h3>Themes</h3>
        <ul>
            <li>Custom background colors</li>
            <li>Accent color customization</li>
            <li>Dark theme support</li>
        </ul>
        
        <h3>Settings</h3>
        <ul>
            <li>Auto-save configuration</li>
            <li>Save interval adjustment</li>
            <li>Theme preferences</li>
            <li>Display options</li>
        </ul>
        """
        
        widget.setHtml(content)
        return widget
    
    def create_formatting_tab(self):
        """Formatting guide"""
        widget = QTextBrowser()
        
        content = """
        <h1>Formatting Guide</h1>
        
        <h2>Column Formatting</h2>
        <p>Make your data visually appealing and easier to read with column formatting.</p>
        
        <h3>How to Format Columns</h3>
        <ol>
            <li>Click the <b>Format Columns</b> button in the toolbar</li>
            <li>Select a column from the dropdown list</li>
            <li>Configure your desired formatting options</li>
            <li>Click <b>Apply</b> to see changes immediately</li>
            <li>Format additional columns as needed</li>
            <li>Click <b>Done</b> when finished</li>
        </ol>
        
        <h3>Font Settings</h3>
        <table border="1" cellpadding="5" cellspacing="0" width="100%">
            <tr>
                <th>Setting</th>
                <th>Options</th>
                <th>Description</th>
            </tr>
            <tr>
                <td><b>Font Name</b></td>
                <td>Calibri, Arial, Times New Roman, etc.</td>
                <td>Choose the typeface for your text</td>
            </tr>
            <tr>
                <td><b>Font Size</b></td>
                <td>8 - 72 points</td>
                <td>Text size in points</td>
            </tr>
            <tr>
                <td><b>Bold</b></td>
                <td>Checkbox</td>
                <td>Make text bold</td>
            </tr>
            <tr>
                <td><b>Italic</b></td>
                <td>Checkbox</td>
                <td>Make text italic</td>
            </tr>
            <tr>
                <td><b>Underline</b></td>
                <td>Checkbox</td>
                <td>Underline text</td>
            </tr>
        </table>
        
        <h3>Color Settings</h3>
        <table border="1" cellpadding="5" cellspacing="0" width="100%">
            <tr>
                <th>Setting</th>
                <th>Description</th>
            </tr>
            <tr>
                <td><b>Text Color</b></td>
                <td>Color of the text - click button to choose from color picker</td>
            </tr>
            <tr>
                <td><b>Background Color</b></td>
                <td>Cell background color - helps highlight important columns</td>
            </tr>
            <tr>
                <td><b>Header Text Color</b></td>
                <td>Color for the column header text</td>
            </tr>
            <tr>
                <td><b>Header Background</b></td>
                <td>Background color for column headers</td>
            </tr>
        </table>
        
        <h3>Alignment Settings</h3>
        <ul>
            <li><b>Horizontal:</b> Left, Center, Right, General</li>
            <li><b>Vertical:</b> Top, Center, Bottom</li>
            <li><b>Wrap Text:</b> Enable to wrap long text within cells</li>
        </ul>
        
        <h3>Column Width</h3>
        <ul>
            <li><b>Auto:</b> Automatically adjust to content width</li>
            <li><b>Custom:</b> Set specific width in pixels (50-500)</li>
        </ul>
        
        <h3>Number Formats</h3>
        
        <h4>General</h4>
        <p>Default format - displays numbers as entered.</p>
        
        <h4>Number</h4>
        <ul>
            <li>Decimal places: 0-10</li>
            <li>Thousands separator option</li>
            <li>Example: 1234.56 → 1,234.56</li>
        </ul>
        
        <h4>Currency</h4>
        <ul>
            <li>Currency symbol ($, €, £, ¥, etc.)</li>
            <li>Decimal places</li>
            <li>Thousands separator</li>
            <li>Example: 1234.5 → $1,234.50</li>
        </ul>
        
        <h4>Accounting</h4>
        <ul>
            <li>Aligns currency symbols</li>
            <li>Shows negative in parentheses</li>
            <li>Example: -1234.5 → ($1,234.50)</li>
        </ul>
        
        <h4>Percentage</h4>
        <ul>
            <li>Displays as percentage</li>
            <li>Decimal places option</li>
            <li>Example: 0.125 → 12.5%</li>
        </ul>
        
        <h4>Date</h4>
        <p>Formats as date: mm/dd/yyyy</p>
        
        <h4>Time</h4>
        <p>Formats as time: h:mm:ss AM/PM</p>
        
        <h4>Scientific</h4>
        <ul>
            <li>Scientific notation</li>
            <li>Example: 1234567 → 1.23E+06</li>
        </ul>
        
        <h4>Text</h4>
        <p>Treats numbers as text (preserves leading zeros, etc.)</p>
        
        <h2>Best Practices</h2>
        
        <h3>Color Usage</h3>
        <ul>
            <li>✅ Use consistent colors for similar data types</li>
            <li>✅ Ensure good contrast between text and background</li>
            <li>✅ Use headers with distinct colors to separate from data</li>
            <li>❌ Avoid too many different colors - keep it professional</li>
        </ul>
        
        <h3>Number Formatting</h3>
        <ul>
            <li>✅ Use currency format for monetary values</li>
            <li>✅ Use percentage for ratios and rates</li>
            <li>✅ Consistent decimal places for numbers in same column</li>
            <li>✅ Use thousands separator for large numbers</li>
        </ul>
        
        <h3>Text Alignment</h3>
        <ul>
            <li>✅ Left-align text data</li>
            <li>✅ Right-align numbers</li>
            <li>✅ Center-align headers</li>
            <li>✅ Use vertical center for better readability</li>
        </ul>
        
        <h2>Visual Feedback</h2>
        <p><b>New in Version 3.1.1:</b> Formatting changes now appear <b>immediately</b> in the table view!</p>
        <ul>
            <li>See your formatting as you apply it</li>
            <li>No need to export to preview</li>
            <li>What you see is what you'll get in Excel</li>
        </ul>
        """
        
        widget.setHtml(content)
        return widget
    
    def create_cloud_sync_tab(self):
        """Cloud Synchronization guide"""
        widget = QTextBrowser()
        
        content = """
        <h1>☁️ Cloud Synchronization Guide</h1>
        
        <h2>Overview</h2>
        <p>Excel Editor Pro now supports seamless integration with major cloud storage services, 
        enabling you to backup, sync, and access your files from anywhere.</p>
        
        <h3>Supported Services</h3>
        <table border="1" cellpadding="8" cellspacing="0" width="100%">
            <tr style="background-color: #f0f0f0;">
                <th>Service</th>
                <th>Features</th>
                <th>Authentication</th>
            </tr>
            <tr>
                <td><b>Google Drive</b></td>
                <td>Upload, Download, Auto-Sync</td>
                <td>OAuth 2.0 / Manual</td>
            </tr>
            <tr>
                <td><b>OneDrive</b></td>
                <td>Upload, Download, Auto-Sync</td>
                <td>OAuth 2.0 / Manual</td>
            </tr>
            <tr>
                <td><b>Dropbox</b></td>
                <td>Upload, Download, Auto-Sync</td>
                <td>OAuth 2.0 / Manual</td>
            </tr>
        </table>
        
        <h2>Getting Started (5 Minutes)</h2>
        
        <h3>Step 1: Connect a Cloud Service</h3>
        <ol>
            <li>Click the <b>☁️ Cloud Sync</b> button or press <b>Ctrl+Shift+U</b></li>
            <li>Go to the <b>Services</b> tab</li>
            <li>Click <b>Connect</b> on your preferred service (Google Drive, OneDrive, or Dropbox)</li>
            <li>Choose authentication method:
                <ul>
                    <li><b>OAuth Authentication (Recommended):</b> Secure browser-based login</li>
                    <li><b>Manual Credentials:</b> Enter API credentials if you have them</li>
                </ul>
            </li>
            <li>Follow the prompts to authorize Excel Editor Pro</li>
            <li>✓ Connected! The service card will show green status</li>
        </ol>
        
        <h3>Step 2: Upload Your First File</h3>
        
        <h4>Quick Upload (Fastest):</h4>
        <ol>
            <li>Open a file in Excel Editor Pro</li>
            <li>Press <b>Ctrl+U</b> (or Menu → Cloud → Quick Upload)</li>
            <li>Select cloud service and destination folder</li>
            <li>Click <b>📤 Upload File</b></li>
            <li>Done! File is now in your cloud storage ✓</li>
        </ol>
        
        <h4>Upload Any File:</h4>
        <ol>
            <li>Open Cloud Sync dialog (<b>Ctrl+Shift+U</b>)</li>
            <li>Go to <b>Upload</b> tab</li>
            <li>Click "Browse..." to select any file</li>
            <li>Or click "Use Current File" for the open file</li>
            <li>Choose service and destination folder</li>
            <li>Configure options (overwrite, compress)</li>
            <li>Click <b>📤 Upload File</b></li>
        </ol>
        
        <h3>Step 3: Enable Auto-Backup (Recommended)</h3>
        <ol>
            <li>In Cloud Sync dialog, go to <b>Auto-Sync</b> tab</li>
            <li>Check <b>"Enable auto-sync for all services"</b></li>
            <li>Select sync interval: <b>"On Save"</b> (recommended)</li>
            <li>Click <b>"Add Folder"</b> to add your work directories</li>
            <li>✓ Your files now backup automatically!</li>
        </ol>
        
        <h2>Cloud Sync Features</h2>
        
        <h3>1. Upload Files</h3>
        <p><b>Upload Tab Features:</b></p>
        <ul>
            <li>Browse and select any file</li>
            <li>Use current open file</li>
            <li>Choose destination service and folder</li>
            <li>Overwrite existing files option</li>
            <li>Compress before upload (saves space & time)</li>
            <li>Real-time progress tracking</li>
            <li>Background upload (non-blocking)</li>
        </ul>
        
        <h3>2. Download Files</h3>
        <p><b>Download Tab Features:</b></p>
        <ul>
            <li>Browse files from any connected service</li>
            <li>Refresh file list with one click</li>
            <li>Select download destination folder</li>
            <li>Real-time progress tracking</li>
            <li>Background download (non-blocking)</li>
            <li>Open downloaded files directly</li>
        </ul>
        
        <h3>3. Auto-Sync</h3>
        <p><b>Automatic Synchronization Options:</b></p>
        <ul>
            <li><b>On Save:</b> Backup every time you save (recommended)</li>
            <li><b>Every 5 minutes:</b> Frequent sync for active work</li>
            <li><b>Every 15 minutes:</b> Balanced sync interval</li>
            <li><b>Every 30 minutes:</b> Less frequent, saves bandwidth</li>
            <li><b>Every hour:</b> Minimal sync for stable files</li>
        </ul>
        
        <p><b>Folder Synchronization:</b></p>
        <ul>
            <li>Add entire folders to sync list</li>
            <li>All files in folder sync automatically</li>
            <li>Remove folders anytime</li>
            <li>Manual "Sync All Now" button</li>
        </ul>
        
        <h3>4. Sync History</h3>
        <p><b>Track All Operations:</b></p>
        <ul>
            <li>Date and time of each operation</li>
            <li>Operation type (Upload, Download, Sync)</li>
            <li>Service used</li>
            <li>File name</li>
            <li>Status (Success or Failed)</li>
            <li>Export history to CSV</li>
            <li>Clear all history</li>
        </ul>
        
        <h2>Keyboard Shortcuts</h2>
        <table border="1" cellpadding="5" cellspacing="0" width="100%">
            <tr style="background-color: #f0f0f0;">
                <th>Shortcut</th>
                <th>Action</th>
            </tr>
            <tr>
                <td><b>Ctrl+Shift+U</b></td>
                <td>Open Cloud Sync dialog</td>
            </tr>
            <tr>
                <td><b>Ctrl+U</b></td>
                <td>Quick Upload current file</td>
            </tr>
        </table>
        
        <h2>Best Practices</h2>
        
        <h3>🔐 Security</h3>
        <ul>
            <li>✓ Use OAuth authentication (most secure)</li>
            <li>✓ Never share API credentials</li>
            <li>✓ Revoke access from cloud service if you stop using app</li>
            <li>✓ Credentials stored locally only (not shared)</li>
        </ul>
        
        <h3>⚡ Performance</h3>
        <ul>
            <li>✓ Enable compression for large files (>5MB)</li>
            <li>✓ Use selective folder sync (only needed folders)</li>
            <li>✓ Schedule syncs during low-activity periods</li>
            <li>✓ Check internet connection for smooth syncing</li>
        </ul>
        
        <h3>📁 Organization</h3>
        <ul>
            <li>✓ Create dedicated cloud folders for Excel files</li>
            <li>✓ Use consistent naming conventions</li>
            <li>✓ Regular cleanup - remove old versions</li>
            <li>✓ Organize by project or client</li>
        </ul>
        
        <h3>💾 Backup Strategy</h3>
        <ul>
            <li>✓ Enable auto-sync for critical files</li>
            <li>✓ Use multiple services for redundancy</li>
            <li>✓ Regular manual backups for important work</li>
            <li>✓ Check sync history periodically</li>
        </ul>
        
        <h2>Troubleshooting</h2>
        
        <h3>Connection Issues</h3>
        <p><b>Problem:</b> Cannot connect to cloud service</p>
        <ul>
            <li>Check internet connection</li>
            <li>Verify account credentials</li>
            <li>Try disconnecting and reconnecting</li>
            <li>Check if service has scheduled maintenance</li>
        </ul>
        
        <h3>Upload/Download Issues</h3>
        <p><b>Problem:</b> Upload fails</p>
        <ul>
            <li>Check available storage space in cloud</li>
            <li>Verify file size doesn't exceed limits</li>
            <li>Ensure stable internet connection</li>
            <li>Try enabling compression</li>
        </ul>
        
        <p><b>Problem:</b> Download fails</p>
        <ul>
            <li>Verify file still exists in cloud</li>
            <li>Check local disk space</li>
            <li>Ensure you have read permissions</li>
            <li>Try refreshing the file list</li>
        </ul>
        
        <h3>Auto-Sync Issues</h3>
        <p><b>Problem:</b> Auto-sync not working</p>
        <ul>
            <li>Verify service is connected (green status)</li>
            <li>Check auto-sync is enabled in settings</li>
            <li>Ensure sync interval is configured</li>
            <li>Review sync history for errors</li>
        </ul>
        
        <h2>Privacy & Data</h2>
        
        <h3>What Excel Editor Pro Can Access</h3>
        <ul>
            <li>Files you explicitly upload or download</li>
            <li>Folders you add to auto-sync</li>
            <li>Basic account info for authentication</li>
        </ul>
        
        <h3>What Excel Editor Pro Cannot Access</h3>
        <ul>
            <li>Files in folders you haven't shared</li>
            <li>Other applications' data</li>
            <li>Your account password</li>
        </ul>
        
        <h3>Data Storage</h3>
        <ul>
            <li>Credentials stored securely on your local machine</li>
            <li>Settings saved in encrypted format</li>
            <li>No data sent to third parties</li>
            <li>Sync history stored locally only</li>
        </ul>
        
        <h2>Quick Reference</h2>
        
        <table border="1" cellpadding="8" cellspacing="0" width="100%">
            <tr style="background-color: #f0f0f0;">
                <th>Task</th>
                <th>Steps</th>
            </tr>
            <tr>
                <td><b>Connect Service</b></td>
                <td>Cloud Sync → Services → Connect → Authenticate</td>
            </tr>
            <tr>
                <td><b>Quick Upload</b></td>
                <td>Ctrl+U → Select service → Upload</td>
            </tr>
            <tr>
                <td><b>Download File</b></td>
                <td>Cloud Sync → Download → Refresh → Select → Download</td>
            </tr>
            <tr>
                <td><b>Enable Auto-Sync</b></td>
                <td>Cloud Sync → Auto-Sync → Enable → Add Folders</td>
            </tr>
            <tr>
                <td><b>View History</b></td>
                <td>Cloud Sync → History tab</td>
            </tr>
            <tr>
                <td><b>Disconnect</b></td>
                <td>Cloud Sync → Services → Disconnect</td>
            </tr>
        </table>
        
        <h2>For More Information</h2>
        <p>See the complete documentation:</p>
        <ul>
            <li><b>CLOUD_SYNC_QUICK_START.md</b> - 5-minute setup guide</li>
            <li><b>CLOUD_SYNC_GUIDE.md</b> - Comprehensive guide with advanced features</li>
        </ul>
        """
        
        widget.setHtml(content)
        return widget
    
    def create_ai_features_tab(self):
        """AI Features guide"""
        widget = QTextBrowser()
        
        content = """
        <h1>🤖 AI Features Guide</h1>
        
        <h2>Overview</h2>
        <p>Excel Editor Pro includes powerful AI features using local machine learning 
        to help you analyze, clean, and understand your data.</p>
        
        <p><b>Note:</b> AI features require additional packages. Install with:</p>
        <pre>pip install scikit-learn scipy</pre>
        
        <h2>Getting Started</h2>
        
        <h3>Opening AI Assistant</h3>
        <ol>
            <li>Open a file with data</li>
            <li>Click the <b>🤖 AI Assistant</b> button</li>
            <li>Or press <b>Ctrl+Shift+A</b></li>
            <li>Or Menu → AI → AI Data Assistant</li>
        </ol>
        
        <h2>AI Features</h2>
        
        <h3>1. Smart Insights 🔍</h3>
        <p>Automatic data analysis and pattern detection:</p>
        <ul>
            <li>Column type detection (email, phone, currency, dates)</li>
            <li>Pattern recognition in your data</li>
            <li>Relationship discovery between columns</li>
            <li>Automatic insights generation</li>
            <li>Statistical anomaly detection</li>
        </ul>
        
        <h3>2. Data Quality Assessment ✅</h3>
        <p>Get a 0-100 quality score for your dataset:</p>
        <ul>
            <li>Completeness check (missing values)</li>
            <li>Consistency analysis</li>
            <li>Validity checking</li>
            <li>Duplicate detection</li>
            <li>Quality recommendations</li>
        </ul>
        
        <h3>3. Smart Data Cleaning 🧹</h3>
        <p>AI-powered cleaning and standardization:</p>
        <ul>
            <li><b>Smart Fill Missing Values:</b> KNN imputation learns from data</li>
            <li><b>Remove Duplicates:</b> Intelligent duplicate detection</li>
            <li><b>Standardize Formats:</b> Consistent formatting</li>
            <li><b>Outlier Removal:</b> Detect and handle anomalies</li>
            <li><b>Text Cleaning:</b> Remove extra spaces, fix casing</li>
        </ul>
        
        <h3>4. Predictive Analytics 🔮</h3>
        <p>Machine learning predictions and forecasting:</p>
        <ul>
            <li><b>Predict Missing Values:</b> Fill gaps intelligently</li>
            <li><b>Trend Forecasting:</b> Predict future values</li>
            <li><b>Classification:</b> Categorize data automatically</li>
            <li><b>Clustering:</b> Find natural groupings (K-Means)</li>
            <li><b>Regression:</b> Predict numeric outcomes</li>
        </ul>
        
        <h3>5. Formula Assistant 📐</h3>
        <p>AI-generated formula suggestions:</p>
        <ul>
            <li>Context-aware recommendations</li>
            <li>Common calculation patterns</li>
            <li>Column combination suggestions</li>
            <li>Aggregation formulas</li>
        </ul>
        
        <h2>Quick AI Operations</h2>
        
        <h3>Quick Analysis (Ctrl+Alt+A)</h3>
        <p>Get instant insights without opening full AI dialog:</p>
        <ol>
            <li>Press <b>Ctrl+Alt+A</b></li>
            <li>View data quality score</li>
            <li>See top 5 insights</li>
            <li>Get column type summary</li>
        </ol>
        
        <h3>Smart Clean Data</h3>
        <p>Quick access to cleaning features:</p>
        <ol>
            <li>Menu → AI → Smart Clean Data</li>
            <li>Opens AI dialog directly to Cleaning tab</li>
            <li>Select cleaning operations</li>
            <li>Apply with one click</li>
        </ol>
        
        <h3>Get AI Suggestions</h3>
        <p>Receive AI recommendations for your data:</p>
        <ol>
            <li>Menu → AI → Get Suggestions</li>
            <li>View missing data alerts</li>
            <li>See calculation suggestions</li>
            <li>Get action recommendations</li>
        </ol>
        
        <h2>AI Workflow Examples</h2>
        
        <h3>Example 1: Clean Messy Data</h3>
        <ol>
            <li>Open file with missing values</li>
            <li>Click <b>🤖 AI Assistant</b></li>
            <li>Go to <b>Cleaning</b> tab</li>
            <li>Select "Smart Fill Missing Values"</li>
            <li>Choose KNN imputation</li>
            <li>Click <b>Clean Data</b></li>
            <li>✓ Missing values filled intelligently!</li>
        </ol>
        
        <h3>Example 2: Find Data Patterns</h3>
        <ol>
            <li>Open dataset</li>
            <li>Press <b>Ctrl+Shift+A</b></li>
            <li>Go to <b>Insights</b> tab</li>
            <li>Click <b>Analyze Data</b></li>
            <li>Review discovered patterns</li>
            <li>Check quality score</li>
            <li>Read AI recommendations</li>
        </ol>
        
        <h3>Example 3: Predict Values</h3>
        <ol>
            <li>Open data with numeric columns</li>
            <li>Open AI Assistant</li>
            <li>Go to <b>Predictions</b> tab</li>
            <li>Select target column</li>
            <li>Choose prediction method</li>
            <li>Click <b>Run Prediction</b></li>
            <li>View and apply results</li>
        </ol>
        
        <h2>AI Best Practices</h2>
        
        <h3>📊 Data Preparation</h3>
        <ul>
            <li>✓ Ensure data has headers</li>
            <li>✓ Have sufficient data rows (50+ recommended)</li>
            <li>✓ Remove extreme outliers first</li>
            <li>✓ Check data types are correct</li>
        </ul>
        
        <h3>🎯 Using Predictions</h3>
        <ul>
            <li>✓ Start with data quality check</li>
            <li>✓ Clean data before predictions</li>
            <li>✓ Review AI suggestions before applying</li>
            <li>✓ Validate results make sense</li>
        </ul>
        
        <h3>⚡ Performance</h3>
        <ul>
            <li>✓ AI works faster on smaller datasets</li>
            <li>✓ Filter data first for focused analysis</li>
            <li>✓ Use Quick Analysis for instant insights</li>
            <li>✓ Full analysis for comprehensive results</li>
        </ul>
        
        <h2>Keyboard Shortcuts</h2>
        <table border="1" cellpadding="5" cellspacing="0" width="100%">
            <tr style="background-color: #f0f0f0;">
                <th>Shortcut</th>
                <th>Action</th>
            </tr>
            <tr>
                <td><b>Ctrl+Shift+A</b></td>
                <td>Open AI Assistant</td>
            </tr>
            <tr>
                <td><b>Ctrl+Alt+A</b></td>
                <td>Quick AI Analysis</td>
            </tr>
        </table>
        
        <h2>For More Information</h2>
        <p>See the complete AI documentation:</p>
        <ul>
            <li><b>AI_QUICK_START.md</b> - Get started in 5 minutes</li>
            <li><b>AI_FEATURES_GUIDE.md</b> - Comprehensive AI guide</li>
        </ul>
        """
        
        widget.setHtml(content)
        return widget
    
    def create_transformation_tab(self):
        """Data transformation guide"""
        widget = QTextBrowser()
        
        content = """
        <h1>Data Transformation Guide</h1>
        
        <h2>Overview</h2>
        <p>The Data Transformation feature helps you clean, modify, and enhance your data with powerful operations.</p>
        
        <h3>How to Access</h3>
        <ol>
            <li>Click the <b>Transform Data</b> button (or press <b>Ctrl+T</b>)</li>
            <li>Choose the transformation type</li>
            <li>Configure transformation settings</li>
            <li>Preview the results</li>
            <li>Apply to your data</li>
        </ol>
        
        <h2>Text Transformations</h2>
        
        <h3>Change Case</h3>
        <ul>
            <li><b>UPPERCASE:</b> Convert all text to capital letters</li>
            <li><b>lowercase:</b> Convert all text to small letters</li>
            <li><b>Title Case:</b> Capitalize First Letter Of Each Word</li>
            <li><b>Sentence case:</b> Capitalize first letter of sentences</li>
        </ul>
        
        <h3>Trim & Clean</h3>
        <ul>
            <li><b>Trim spaces:</b> Remove leading and trailing whitespace</li>
            <li><b>Remove extra spaces:</b> Replace multiple spaces with single space</li>
            <li><b>Remove special characters:</b> Clean punctuation and symbols</li>
        </ul>
        
        <h3>Find & Replace</h3>
        <ul>
            <li>Find specific text</li>
            <li>Replace with new text</li>
            <li>Case-sensitive or case-insensitive</li>
            <li>Apply to entire column or selected cells</li>
        </ul>
        
        <h3>Split Text</h3>
        <ul>
            <li>Split by delimiter (comma, space, custom)</li>
            <li>Extract parts into new columns</li>
            <li>Example: "John Doe" → "John" | "Doe"</li>
        </ul>
        
        <h2>Number Transformations</h2>
        
        <h3>Mathematical Operations</h3>
        <ul>
            <li><b>Add:</b> Add a constant to all values</li>
            <li><b>Subtract:</b> Subtract a constant from all values</li>
            <li><b>Multiply:</b> Multiply all values by a constant</li>
            <li><b>Divide:</b> Divide all values by a constant</li>
        </ul>
        
        <h3>Rounding</h3>
        <ul>
            <li><b>Round:</b> Round to nearest integer or decimal place</li>
            <li><b>Round up:</b> Always round up (ceiling)</li>
            <li><b>Round down:</b> Always round down (floor)</li>
            <li>Specify decimal places</li>
        </ul>
        
        <h3>Formatting</h3>
        <ul>
            <li>Add thousand separators</li>
            <li>Set decimal places</li>
            <li>Convert to percentage</li>
            <li>Add currency symbols</li>
        </ul>
        
        <h2>Date Transformations</h2>
        
        <h3>Date Parsing</h3>
        <ul>
            <li>Convert text to dates</li>
            <li>Multiple format support</li>
            <li>Auto-detection of date formats</li>
        </ul>
        
        <h3>Date Formatting</h3>
        <ul>
            <li>Standard formats (MM/DD/YYYY, DD-MM-YYYY, etc.)</li>
            <li>Custom formats</li>
            <li>Extract components (year, month, day)</li>
        </ul>
        
        <h3>Date Calculations</h3>
        <ul>
            <li>Add/subtract days, months, years</li>
            <li>Calculate age from birthdate</li>
            <li>Calculate date differences</li>
        </ul>
        
        <h2>Data Cleaning</h2>
        
        <h3>Remove Duplicates</h3>
        <ul>
            <li>Identify duplicate rows</li>
            <li>Keep first or last occurrence</li>
            <li>Based on specific columns or all columns</li>
        </ul>
        
        <h3>Handle Missing Values</h3>
        <ul>
            <li><b>Fill with value:</b> Replace blanks with specific value</li>
            <li><b>Fill forward:</b> Use previous valid value</li>
            <li><b>Fill backward:</b> Use next valid value</li>
            <li><b>Remove rows:</b> Delete rows with missing values</li>
        </ul>
        
        <h3>Data Type Conversion</h3>
        <ul>
            <li>Convert text to numbers</li>
            <li>Convert numbers to text</li>
            <li>Parse dates from text</li>
            <li>Handle conversion errors gracefully</li>
        </ul>
        
        <h2>Column Operations</h2>
        
        <h3>Merge Columns</h3>
        <ul>
            <li>Combine multiple columns into one</li>
            <li>Custom separator</li>
            <li>Example: First Name + Last Name → Full Name</li>
        </ul>
        
        <h3>Extract Data</h3>
        <ul>
            <li>Extract by position (first N characters, last N characters)</li>
            <li>Extract by pattern (regex support)</li>
            <li>Extract numbers from text</li>
            <li>Extract emails, phone numbers, etc.</li>
        </ul>
        
        <h2>Advanced Operations</h2>
        
        <h3>Conditional Transformations</h3>
        <ul>
            <li>Apply transformation only if condition is met</li>
            <li>Example: Convert to uppercase only if length > 5</li>
        </ul>
        
        <h3>Bulk Operations</h3>
        <ul>
            <li>Apply same transformation to multiple columns</li>
            <li>Chain multiple transformations</li>
            <li>Save transformation recipes for reuse</li>
        </ul>
        
        <h2>Best Practices</h2>
        
        <h3>Before Transforming</h3>
        <ul>
            <li>✅ Always save your file first</li>
            <li>✅ Make a backup of important data</li>
            <li>✅ Preview transformations before applying</li>
            <li>✅ Test on a small sample first</li>
        </ul>
        
        <h3>Common Workflows</h3>
        <ol>
            <li><b>Data Import Cleanup:</b>
                <ul>
                    <li>Trim whitespace</li>
                    <li>Standardize case</li>
                    <li>Remove duplicates</li>
                    <li>Handle missing values</li>
                </ul>
            </li>
            <li><b>Name Standardization:</b>
                <ul>
                    <li>Convert to Title Case</li>
                    <li>Trim spaces</li>
                    <li>Split into First/Last name if needed</li>
                </ul>
            </li>
            <li><b>Number Formatting:</b>
                <ul>
                    <li>Round to consistent decimal places</li>
                    <li>Add thousand separators</li>
                    <li>Convert to currency format</li>
                </ul>
            </li>
        </ol>
        
        <h2>Tips & Tricks</h2>
        <ul>
            <li>💡 Use Preview to see results before applying</li>
            <li>💡 Transformations can be undone with Ctrl+Z</li>
            <li>💡 Chain transformations for complex operations</li>
            <li>💡 Save often when doing multiple transformations</li>
        </ul>
        """
        
        widget.setHtml(content)
        return widget
    
    def create_troubleshooting_tab(self):
        """Troubleshooting guide"""
        widget = QTextBrowser()
        
        content = """
        <h1>Troubleshooting Guide</h1>
        
        <h2>Common Issues & Solutions</h2>
        
        <h3>File Won't Open</h3>
        <p><b>Problem:</b> Error message when trying to open a file</p>
        <p><b>Solutions:</b></p>
        <ul>
            <li>Ensure the file is not open in another program (Excel, etc.)</li>
            <li>Check file is not corrupted - try opening in Excel first</li>
            <li>Verify file extension matches actual format (.xlsx should be Excel)</li>
            <li>For large files, ensure sufficient memory is available</li>
        </ul>
        
        <h3>Can't Save File</h3>
        <p><b>Problem:</b> Error when trying to save</p>
        <p><b>Solutions:</b></p>
        <ul>
            <li>Check you have write permissions to the folder</li>
            <li>Ensure the file is not marked as read-only</li>
            <li>Close the file in other programs if open</li>
            <li>Try "Save As" to a different location</li>
            <li>Ensure disk has sufficient space</li>
        </ul>
        
        <h3>Formatting Not Appearing</h3>
        <p><b>Problem:</b> Applied formatting doesn't show</p>
        <p><b>Solutions:</b></p>
        <ul>
            <li>Ensure you're using version 3.1.1 or later</li>
            <li>Click the table to refresh view</li>
            <li>Try closing and reopening the file</li>
            <li>Check Format Columns dialog shows your settings</li>
        </ul>
        
        <h3>Data Not Showing</h3>
        <p><b>Problem:</b> Table appears empty or incomplete</p>
        <p><b>Solutions:</b></p>
        <ul>
            <li>Check if a filter is active - click "Clear Filter"</li>
            <li>For large files (>1000 rows), only first 1000 display</li>
            <li>Scroll horizontally and vertically to see all data</li>
            <li>Try refreshing with F5</li>
        </ul>
        
        <h3>Application Crashes</h3>
        <p><b>Problem:</b> Program closes unexpectedly</p>
        <p><b>Solutions:</b></p>
        <ul>
            <li>Ensure all required libraries are installed:
                <ul>
                    <li>PyQt5</li>
                    <li>pandas</li>
                    <li>openpyxl</li>
                    <li>python-docx</li>
                </ul>
            </li>
            <li>Try with a smaller file to isolate issue</li>
            <li>Check available system memory</li>
            <li>Enable auto-save to prevent data loss</li>
        </ul>
        
        <h3>Slow Performance</h3>
        <p><b>Problem:</b> Application is slow or laggy</p>
        <p><b>Solutions:</b></p>
        <ul>
            <li>Close other applications to free memory</li>
            <li>For large files, use filters to work with subsets</li>
            <li>Disable auto-save if working with very large files</li>
            <li>Consider splitting large files into smaller ones</li>
        </ul>
        
        <h3>Multi-Sheet Files</h3>
        <p><b>Problem:</b> Can't see all sheets in Excel file</p>
        <p><b>Solutions:</b></p>
        <ul>
            <li>When opening, dialog shows all sheets - select one</li>
            <li>Currently only one sheet can be edited at a time</li>
            <li>To edit another sheet, save and reopen file, select different sheet</li>
        </ul>
        
        <h3>Copy/Paste Not Working</h3>
        <p><b>Problem:</b> Can't copy or paste data</p>
        <p><b>Solutions:</b></p>
        <ul>
            <li>Use double-click to edit cells, then Ctrl+C/Ctrl+V</li>
            <li>Right-click for context menu options</li>
            <li>For bulk operations, use Import/Export features</li>
        </ul>
        
        <h2>Error Messages</h2>
        
        <h3>"No module named 'openpyxl'"</h3>
        <p><b>Solution:</b> Install required package:</p>
        <code>pip install openpyxl</code>
        
        <h3>"Permission denied"</h3>
        <p><b>Solution:</b> File is open elsewhere or you lack permissions</p>
        <ul>
            <li>Close file in other programs</li>
            <li>Run application with appropriate permissions</li>
            <li>Try saving to a different location</li>
        </ul>
        
        <h3>"Memory error"</h3>
        <p><b>Solution:</b> File too large for available memory</p>
        <ul>
            <li>Close other applications</li>
            <li>Split file into smaller parts</li>
            <li>Increase system RAM if possible</li>
        </ul>
        
        <h2>Data Issues</h2>
        
        <h3>Numbers Saved as Text</h3>
        <p><b>Problem:</b> Numbers treated as text</p>
        <p><b>Solutions:</b></p>
        <ul>
            <li>Use Data Transformation → Convert to Number</li>
            <li>Apply Number format in Column Formatting</li>
            <li>Remove any leading/trailing spaces with Trim</li>
        </ul>
        
        <h3>Dates Not Recognized</h3>
        <p><b>Problem:</b> Dates showing as text or numbers</p>
        <p><b>Solutions:</b></p>
        <ul>
            <li>Use Data Transformation → Parse Dates</li>
            <li>Apply Date format in Column Formatting</li>
            <li>Ensure consistent date format in data</li>
        </ul>
        
        <h3>Lost Decimal Places</h3>
        <p><b>Problem:</b> Decimal values rounded unexpectedly</p>
        <p><b>Solutions:</b></p>
        <ul>
            <li>Check Column Formatting → Number Format</li>
            <li>Set appropriate decimal places</li>
            <li>Ensure column is formatted as Number, not General</li>
        </ul>
        
        <h2>Cloud Sync Issues</h2>
        
        <h3>Can't Connect to Cloud Service</h3>
        <p><b>Problem:</b> Connection to Google Drive/OneDrive/Dropbox fails</p>
        <p><b>Solutions:</b></p>
        <ul>
            <li>Check your internet connection</li>
            <li>Verify your account credentials are correct</li>
            <li>Try disconnecting and reconnecting the service</li>
            <li>Check if the cloud service is experiencing downtime</li>
            <li>Ensure OAuth popup isn't blocked by browser</li>
            <li>Try manual credential entry as alternative</li>
        </ul>
        
        <h3>Upload Fails</h3>
        <p><b>Problem:</b> File upload to cloud doesn't complete</p>
        <p><b>Solutions:</b></p>
        <ul>
            <li>Check available storage space in cloud account</li>
            <li>Verify file size doesn't exceed service limits</li>
            <li>Ensure stable internet connection</li>
            <li>Try enabling compression for large files</li>
            <li>Check if file is locked or in use</li>
            <li>Review sync history for error details</li>
        </ul>
        
        <h3>Download Fails</h3>
        <p><b>Problem:</b> Can't download files from cloud</p>
        <p><b>Solutions:</b></p>
        <ul>
            <li>Verify file still exists in cloud storage</li>
            <li>Check local disk space available</li>
            <li>Ensure you have read permissions for the file</li>
            <li>Try refreshing the file list</li>
            <li>Check internet connection stability</li>
        </ul>
        
        <h3>Auto-Sync Not Working</h3>
        <p><b>Problem:</b> Files not syncing automatically</p>
        <p><b>Solutions:</b></p>
        <ul>
            <li>Verify service is connected (check green status in Services tab)</li>
            <li>Ensure auto-sync is enabled in settings</li>
            <li>Check sync interval is configured</li>
            <li>Review sync history for error messages</li>
            <li>Verify folders added to sync list exist</li>
            <li>Check if files are locked or in use</li>
        </ul>
        
        <h3>Authentication Expired</h3>
        <p><b>Problem:</b> "Authentication failed" error</p>
        <p><b>Solutions:</b></p>
        <ul>
            <li>Disconnect and reconnect the service</li>
            <li>Re-authenticate using OAuth</li>
            <li>Check if you changed your cloud account password</li>
            <li>Verify API credentials if using manual method</li>
        </ul>
        
        <h3>"Cloud dependencies not installed"</h3>
        <p><b>Solution:</b> Install cloud sync packages:</p>
        <code>pip install google-auth google-auth-oauthlib google-api-python-client msal dropbox</code>
        
        <h2>AI Features Issues</h2>
        
        <h3>AI Features Not Available</h3>
        <p><b>Problem:</b> "AI features require additional packages" message</p>
        <p><b>Solution:</b> Install AI dependencies:</p>
        <code>pip install scikit-learn scipy</code>
        
        <h3>AI Analysis Slow</h3>
        <p><b>Problem:</b> AI operations taking too long</p>
        <p><b>Solutions:</b></p>
        <ul>
            <li>Use filters to reduce dataset size first</li>
            <li>Try Quick Analysis for faster results</li>
            <li>Close other applications to free memory</li>
            <li>Consider working with data sample for testing</li>
        </ul>
        
        <h3>Prediction Results Don't Make Sense</h3>
        <p><b>Problem:</b> AI predictions seem incorrect</p>
        <p><b>Solutions:</b></p>
        <ul>
            <li>Run data quality check first</li>
            <li>Clean data before predictions</li>
            <li>Ensure sufficient data rows (50+ recommended)</li>
            <li>Check for and remove extreme outliers</li>
            <li>Verify data types are correct</li>
        </ul>
        
        <h2>Getting More Help</h2>
        
        <h3>Feature Requests</h3>
        <p>Have an idea for a new feature? We'd love to hear it!</p>
        
        <h3>Bug Reports</h3>
        <p>Found a bug? Please report with:</p>
        <ul>
            <li>Steps to reproduce the issue</li>
            <li>Error message (if any)</li>
            <li>File size and type</li>
            <li>Application version</li>
        </ul>
        
        <h3>Community</h3>
        <p>Check online forums and communities for:</p>
        <ul>
            <li>Tips and tricks from other users</li>
            <li>Workflow ideas</li>
            <li>Advanced techniques</li>
        </ul>
        """
        
        widget.setHtml(content)
        return widget
    
    def create_about_tab(self):
        """About information"""
        widget = QTextBrowser()
        
        content = """
        <h1>About Excel Editor Pro</h1>
        
        <h2>Application Information</h2>
        <table border="0" cellpadding="5" width="100%">
            <tr>
                <td width="30%"><b>Application:</b></td>
                <td>Excel Editor Pro</td>
            </tr>
            <tr>
                <td><b>Version:</b></td>
                <td>3.4.0 - Cloud Sync Edition</td>
            </tr>
            <tr>
                <td><b>Release Date:</b></td>
                <td>February 8, 2026</td>
            </tr>
            <tr>
                <td><b>Type:</b></td>
                <td>Spreadsheet Editor & Data Management Tool</td>
            </tr>
        </table>
        
        <h2>Description</h2>
        <p>Excel Editor Pro is a powerful, user-friendly application for editing Excel and CSV files. 
        It combines ease of use with professional features, making data management accessible to everyone.</p>
        
        <h2>Key Features</h2>
        <ul>
            <li>☁️ <b>Cloud Synchronization (NEW!):</b> Google Drive, OneDrive, Dropbox integration</li>
            <li>🤖 <b>AI-Powered Operations (NEW!):</b> Smart insights, predictions, and data cleaning</li>
            <li>✨ Edit Excel (.xlsx, .xls) and CSV files</li>
            <li>🎨 Rich formatting options with instant preview</li>
            <li>🔄 Powerful data transformation tools</li>
            <li>📊 Built-in statistics and analysis</li>
            <li>💾 Auto-save functionality</li>
            <li>🔄 Auto-sync to cloud storage</li>
            <li>🎯 Filtering and search capabilities</li>
            <li>📝 Export to Word documents</li>
            <li>⚡ Fast performance with large files</li>
            <li>🎨 Customizable themes</li>
            <li>⌨️ Comprehensive keyboard shortcuts</li>
            <li>↶↷ Undo/Redo system</li>
            <li>📚 Version history tracking</li>
        </ul>
        
        <h2>What's New in Version 3.4 - Cloud Sync Edition</h2>
        <ul>
            <li>☁️ <b>Cloud Storage Integration:</b> Connect to Google Drive, OneDrive, and Dropbox</li>
            <li>📤 <b>Quick Upload:</b> One-click upload with Ctrl+U</li>
            <li>📥 <b>Easy Download:</b> Browse and download files from cloud</li>
            <li>🔄 <b>Auto-Sync:</b> Automatic backup on save or scheduled intervals</li>
            <li>📁 <b>Folder Sync:</b> Sync entire directories automatically</li>
            <li>📊 <b>Sync History:</b> Track all upload/download operations</li>
            <li>🔐 <b>Secure OAuth:</b> Industry-standard authentication</li>
            <li>💾 <b>Multi-Cloud Support:</b> Use multiple services simultaneously</li>
        </ul>
        
        <h2>Version 3.3 - AI Enhanced Edition</h2>
        <ul>
            <li>🤖 <b>AI Data Assistant:</b> Machine learning powered analysis</li>
            <li>🔍 <b>Smart Insights:</b> Automatic pattern detection</li>
            <li>✅ <b>Data Quality Scoring:</b> 0-100 health assessment</li>
            <li>🧹 <b>Smart Cleaning:</b> AI-powered data cleaning</li>
            <li>🔮 <b>Predictions:</b> ML-based forecasting and clustering</li>
            <li>📐 <b>Formula Assistant:</b> AI-generated suggestions</li>
        </ul>
        
        <h2>Technology Stack</h2>
        <table border="1" cellpadding="5" cellspacing="0" width="100%">
            <tr>
                <th>Component</th>
                <th>Technology</th>
            </tr>
            <tr>
                <td>User Interface</td>
                <td>PyQt5</td>
            </tr>
            <tr>
                <td>Data Processing</td>
                <td>pandas, NumPy</td>
            </tr>
            <tr>
                <td>Excel Support</td>
                <td>openpyxl</td>
            </tr>
            <tr>
                <td>Word Export</td>
                <td>python-docx</td>
            </tr>
            <tr>
                <td>AI & Machine Learning</td>
                <td>scikit-learn, scipy</td>
            </tr>
            <tr>
                <td>Cloud Integration</td>
                <td>Google Drive API, OneDrive API, Dropbox SDK</td>
            </tr>
            <tr>
                <td>Language</td>
                <td>Python 3.7+</td>
            </tr>
        </table>
        
        <h2>System Requirements</h2>
        <h3>Minimum Requirements</h3>
        <ul>
            <li>Operating System: Windows 7/10/11, macOS 10.13+, Linux</li>
            <li>Python: 3.7 or higher</li>
            <li>RAM: 2 GB</li>
            <li>Disk Space: 100 MB</li>
        </ul>
        
        <h3>Recommended Requirements</h3>
        <ul>
            <li>Operating System: Windows 10/11, macOS 11+, Ubuntu 20.04+</li>
            <li>Python: 3.9 or higher</li>
            <li>RAM: 4 GB or more</li>
            <li>Disk Space: 500 MB</li>
        </ul>
        
        <h2>License & Terms</h2>
        <p>Excel Editor Pro is provided as-is for data editing and management purposes.</p>
        
        <h2>Credits</h2>
        <p>This application was built using several excellent open-source libraries:</p>
        <ul>
            <li><b>PyQt5:</b> Cross-platform GUI framework</li>
            <li><b>pandas:</b> Data manipulation and analysis</li>
            <li><b>openpyxl:</b> Excel file reading and writing</li>
            <li><b>python-docx:</b> Word document creation</li>
            <li><b>NumPy:</b> Numerical computing</li>
        </ul>
        
        <h2>Contact & Support</h2>
        <p>For questions, suggestions, or issues, please reach out through:</p>
        <ul>
            <li>📧 Email support</li>
            <li>💬 Community forums</li>
            <li>🐛 Bug tracker</li>
        </ul>
        
        <hr>
        <p align="center"><i>Thank you for using Excel Editor Pro!</i></p>
        <p align="center">© 2026 Excel Editor Pro. All rights reserved.</p>
        """
        
        widget.setHtml(content)
        return widget
