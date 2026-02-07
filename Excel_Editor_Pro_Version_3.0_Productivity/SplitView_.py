"""
Split View Module
Allows viewing different parts of large spreadsheets simultaneously
"""

from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QPushButton,
                             QTableWidget, QTableWidgetItem, QLabel, QSplitter,
                             QGroupBox, QRadioButton, QMessageBox, QHeaderView)
from PyQt5.QtCore import Qt, QSettings
from PyQt5.QtGui import QFont


class SplitViewManager:
    """Manages split view functionality"""
    
    def __init__(self, parent):
        self.parent = parent
        self.settings = QSettings("DataTools", "ExcelEditor")
        self.split_active = False
        self.split_window = None
    
    def create_split_view(self, orientation='horizontal'):
        """Create split view window"""
        try:
            if self.parent.df is None:
                QMessageBox.information(
                    self.parent,
                    "No Data",
                    "Please load a file first to use split view."
                )
                return
            
            # Close existing split view if any
            if self.split_window:
                self.split_window.close()
            
            # Create split view dialog
            self.split_window = SplitViewDialog(self.parent, orientation)
            self.split_window.show()
            self.split_active = True
            
            self.parent.statusBar().showMessage("Split view activated", 2000)
            
        except Exception as e:
            QMessageBox.critical(
                self.parent,
                "Error",
                f"Failed to create split view:\n{str(e)}"
            )
    
    def close_split_view(self):
        """Close split view"""
        if self.split_window:
            self.split_window.close()
            self.split_window = None
        self.split_active = False
    
    def is_active(self):
        """Check if split view is active"""
        return self.split_active and self.split_window is not None


class SplitViewDialog(QDialog):
    """Dialog for split view with two synchronized table views"""
    
    def __init__(self, parent, orientation='horizontal'):
        super().__init__(parent)
        self.parent = parent
        self.orientation = orientation
        self.init_ui()
        self.populate_tables()
    
    def init_ui(self):
        self.setWindowTitle("Split View")
        self.setGeometry(100, 100, 1400, 800)
        
        layout = QVBoxLayout()
        
        # Header with controls
        header_layout = QHBoxLayout()
        
        header_layout.addWidget(QLabel("📊 Split View - View different sections simultaneously"))
        header_layout.addStretch()
        
        # Orientation toggle
        orientation_group = QGroupBox("Split Orientation")
        orientation_layout = QHBoxLayout()
        
        self.horizontal_radio = QRadioButton("Horizontal")
        self.horizontal_radio.setChecked(self.orientation == 'horizontal')
        self.horizontal_radio.toggled.connect(self.change_orientation)
        orientation_layout.addWidget(self.horizontal_radio)
        
        self.vertical_radio = QRadioButton("Vertical")
        self.vertical_radio.setChecked(self.orientation == 'vertical')
        orientation_layout.addWidget(self.vertical_radio)
        
        orientation_group.setLayout(orientation_layout)
        header_layout.addWidget(orientation_group)
        
        # Sync checkbox
        from PyQt5.QtWidgets import QCheckBox
        self.sync_scroll_checkbox = QCheckBox("Sync Scrolling")
        self.sync_scroll_checkbox.setChecked(False)
        header_layout.addWidget(self.sync_scroll_checkbox)
        
        layout.addLayout(header_layout)
        
        # Create splitter
        if self.orientation == 'horizontal':
            self.splitter = QSplitter(Qt.Horizontal)
        else:
            self.splitter = QSplitter(Qt.Vertical)
        
        # Create two table widgets
        self.table1 = self.create_table_widget("View 1")
        self.table2 = self.create_table_widget("View 2")
        
        # Add to splitter
        view1_widget = self.create_view_panel("View 1", self.table1)
        view2_widget = self.create_view_panel("View 2", self.table2)
        
        self.splitter.addWidget(view1_widget)
        self.splitter.addWidget(view2_widget)
        
        # Set equal sizes
        self.splitter.setSizes([700, 700])
        
        layout.addWidget(self.splitter)
        
        # Navigation controls
        nav_layout = QHBoxLayout()
        
        # View 1 controls
        nav_layout.addWidget(QLabel("View 1:"))
        
        self.goto_top1_btn = QPushButton("↑ Top")
        self.goto_top1_btn.clicked.connect(lambda: self.goto_position(self.table1, 'top'))
        nav_layout.addWidget(self.goto_top1_btn)
        
        self.goto_bottom1_btn = QPushButton("↓ Bottom")
        self.goto_bottom1_btn.clicked.connect(lambda: self.goto_position(self.table1, 'bottom'))
        nav_layout.addWidget(self.goto_bottom1_btn)
        
        nav_layout.addSpacing(20)
        nav_layout.addWidget(QLabel("│"))
        nav_layout.addSpacing(20)
        
        # View 2 controls
        nav_layout.addWidget(QLabel("View 2:"))
        
        self.goto_top2_btn = QPushButton("↑ Top")
        self.goto_top2_btn.clicked.connect(lambda: self.goto_position(self.table2, 'top'))
        nav_layout.addWidget(self.goto_top2_btn)
        
        self.goto_bottom2_btn = QPushButton("↓ Bottom")
        self.goto_bottom2_btn.clicked.connect(lambda: self.goto_position(self.table2, 'bottom'))
        nav_layout.addWidget(self.goto_bottom2_btn)
        
        nav_layout.addStretch()
        
        # Close button
        close_btn = QPushButton("Close Split View")
        close_btn.clicked.connect(self.close)
        nav_layout.addWidget(close_btn)
        
        layout.addLayout(nav_layout)
        
        self.setLayout(layout)
        
        # Setup scroll synchronization if enabled
        if self.sync_scroll_checkbox.isChecked():
            self.setup_scroll_sync()
    
    def create_table_widget(self, name):
        """Create a table widget"""
        table = QTableWidget()
        table.setSortingEnabled(False)
        table.setAlternatingRowColors(True)
        table.setSelectionBehavior(QTableWidget.SelectRows)
        table.setVerticalScrollMode(QTableWidget.ScrollPerPixel)
        table.setHorizontalScrollMode(QTableWidget.ScrollPerPixel)
        table.setObjectName(name)
        
        return table
    
    def create_view_panel(self, title, table):
        """Create a panel with title and table"""
        from PyQt5.QtWidgets import QWidget
        widget = QWidget()
        layout = QVBoxLayout()
        layout.setContentsMargins(5, 5, 5, 5)
        
        title_label = QLabel(f"<b>{title}</b>")
        layout.addWidget(title_label)
        
        layout.addWidget(table)
        
        widget.setLayout(layout)
        return widget
    
    def populate_tables(self):
        """Populate both tables with data"""
        try:
            df = self.parent.df
            
            if df is None:
                return
            
            # Populate both tables
            for table in [self.table1, self.table2]:
                table.setRowCount(len(df))
                table.setColumnCount(len(df.columns))
                table.setHorizontalHeaderLabels(df.columns.astype(str))
                
                # Fill data
                for i in range(len(df)):
                    for j, col in enumerate(df.columns):
                        value = str(df.iloc[i, j])
                        item = QTableWidgetItem(value)
                        table.setItem(i, j, item)
                
                # Auto-resize columns
                table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
                table.resizeColumnsToContents()
            
            # Set different initial positions for variety
            # View 1 starts at top
            self.table1.scrollToTop()
            
            # View 2 starts at middle (or bottom if short)
            if len(df) > 20:
                self.table2.scrollToItem(
                    self.table2.item(len(df) // 2, 0),
                    QTableWidget.PositionAtTop
                )
            else:
                self.table2.scrollToBottom()
            
        except Exception as e:
            print(f"Error populating split view tables: {e}")
            QMessageBox.critical(
                self,
                "Error",
                f"Failed to populate tables:\n{str(e)}"
            )
    
    def change_orientation(self):
        """Change split orientation"""
        try:
            if self.horizontal_radio.isChecked():
                new_orientation = 'horizontal'
            else:
                new_orientation = 'vertical'
            
            if new_orientation != self.orientation:
                self.orientation = new_orientation
                
                # Recreate splitter with new orientation
                old_splitter = self.splitter
                
                if self.orientation == 'horizontal':
                    self.splitter = QSplitter(Qt.Horizontal)
                else:
                    self.splitter = QSplitter(Qt.Vertical)
                
                # Move widgets to new splitter
                view1 = old_splitter.widget(0)
                view2 = old_splitter.widget(1)
                
                self.splitter.addWidget(view1)
                self.splitter.addWidget(view2)
                self.splitter.setSizes([700, 700])
                
                # Replace in layout
                self.layout().replaceWidget(old_splitter, self.splitter)
                old_splitter.deleteLater()
                
        except Exception as e:
            print(f"Error changing orientation: {e}")
    
    def setup_scroll_sync(self):
        """Setup synchronized scrolling between views"""
        try:
            if self.sync_scroll_checkbox.isChecked():
                # Connect scroll bars
                self.table1.verticalScrollBar().valueChanged.connect(
                    lambda value: self.table2.verticalScrollBar().setValue(value)
                )
                self.table2.verticalScrollBar().valueChanged.connect(
                    lambda value: self.table1.verticalScrollBar().setValue(value)
                )
                
                self.table1.horizontalScrollBar().valueChanged.connect(
                    lambda value: self.table2.horizontalScrollBar().setValue(value)
                )
                self.table2.horizontalScrollBar().valueChanged.connect(
                    lambda value: self.table1.horizontalScrollBar().setValue(value)
                )
            else:
                # Disconnect if unchecked
                try:
                    self.table1.verticalScrollBar().valueChanged.disconnect()
                    self.table2.verticalScrollBar().valueChanged.disconnect()
                    self.table1.horizontalScrollBar().valueChanged.disconnect()
                    self.table2.horizontalScrollBar().valueChanged.disconnect()
                except:
                    pass
        except Exception as e:
            print(f"Error setting up scroll sync: {e}")
    
    def goto_position(self, table, position):
        """Navigate to a specific position in a table"""
        try:
            if position == 'top':
                table.scrollToTop()
            elif position == 'bottom':
                table.scrollToBottom()
            elif position == 'middle':
                middle_row = table.rowCount() // 2
                table.scrollToItem(
                    table.item(middle_row, 0),
                    QTableWidget.PositionAtTop
                )
        except Exception as e:
            print(f"Error navigating to position: {e}")
    
    def closeEvent(self, event):
        """Handle window close"""
        if hasattr(self.parent, 'split_view_manager'):
            self.parent.split_view_manager.split_active = False
            self.parent.split_view_manager.split_window = None
        event.accept()


class SplitViewOptionsDialog(QDialog):
    """Dialog for split view options"""
    
    def __init__(self, parent, split_view_manager):
        super().__init__(parent)
        self.parent = parent
        self.split_view_manager = split_view_manager
        self.init_ui()
    
    def init_ui(self):
        self.setWindowTitle("Split View Options")
        self.setGeometry(300, 300, 350, 200)
        
        layout = QVBoxLayout()
        
        layout.addWidget(QLabel("Choose split view orientation:"))
        
        # Orientation options
        orientation_group = QGroupBox("Orientation")
        orientation_layout = QVBoxLayout()
        
        self.horizontal_radio = QRadioButton("Horizontal (Side by Side)")
        self.horizontal_radio.setChecked(True)
        orientation_layout.addWidget(self.horizontal_radio)
        
        self.vertical_radio = QRadioButton("Vertical (Top and Bottom)")
        orientation_layout.addWidget(self.vertical_radio)
        
        orientation_group.setLayout(orientation_layout)
        layout.addWidget(orientation_group)
        
        layout.addStretch()
        
        # Buttons
        button_layout = QHBoxLayout()
        
        create_btn = QPushButton("Create Split View")
        create_btn.clicked.connect(self.create_split)
        button_layout.addWidget(create_btn)
        
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        button_layout.addWidget(cancel_btn)
        
        layout.addLayout(button_layout)
        
        self.setLayout(layout)
    
    def create_split(self):
        """Create split view with selected orientation"""
        orientation = 'horizontal' if self.horizontal_radio.isChecked() else 'vertical'
        self.split_view_manager.create_split_view(orientation)
        self.accept()
